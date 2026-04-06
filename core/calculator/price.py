"""
Модуль расчёта цены (стоимости в рублях).

Формула: цена = сумма по маршрутам (км * вместимость * коэффициент)
- Вместимость из routes_catalog.routes
- Коэффициент из routes_catalog.coefficients по кварталу
"""
import logging
import time
from calendar import monthrange
from datetime import date
from decimal import Decimal
from typing import Optional

import psycopg2
from psycopg2.extras import RealDictCursor
from psycopg2 import pool as pg_pool

from core.config import settings

logger = logging.getLogger(__name__)

# TTL кеша в секундах (1 час)
CACHE_TTL_SECONDS = 3600

# Кэш для данных из внешней БД
_coefficients_cache: dict[tuple[str, str], dict[str, Decimal]] = {}
_capacities_cache: dict[str, dict[str, int]] = {}
_coefficients_loaded: set[str] = set()  # Контракты для которых загружены коэффициенты
_cache_timestamps: dict[str, float] = {}  # Время загрузки данных по ключу

# Пул подключений к внешней БД
_connection_pool: Optional[pg_pool.SimpleConnectionPool] = None


def _get_connection_pool() -> Optional[pg_pool.SimpleConnectionPool]:
    """Возвращает пул подключений к внешней БД (создаёт при первом вызове)."""
    global _connection_pool
    if _connection_pool is not None:
        return _connection_pool

    if not settings.external_db_url:
        return None

    try:
        _connection_pool = pg_pool.SimpleConnectionPool(
            minconn=1,
            maxconn=5,
            dsn=settings.external_db_url,
            connect_timeout=5,
        )
        logger.info("Пул подключений к внешней БД создан")
        return _connection_pool
    except Exception as e:
        logger.error("Ошибка создания пула подключений: %s", e)
        return None


def _is_cache_valid(cache_key: str) -> bool:
    """Проверяет, не истёк ли TTL кеша."""
    ts = _cache_timestamps.get(cache_key)
    if ts is None:
        return False
    return (time.monotonic() - ts) < CACHE_TTL_SECONDS


def invalidate_cache(contract: str | None = None) -> None:
    """
    Принудительно сбрасывает кеш.

    Args:
        contract: Номер контракта (None = сбросить весь кеш)
    """
    if contract is None:
        _coefficients_cache.clear()
        _capacities_cache.clear()
        _coefficients_loaded.clear()
        _cache_timestamps.clear()
        logger.info("Весь кеш цен сброшен")
    else:
        _coefficients_loaded.discard(contract)
        _capacities_cache.pop(contract, None)
        keys_to_remove = [k for k in _coefficients_cache if k[1] == contract]
        for k in keys_to_remove:
            _coefficients_cache.pop(k, None)
        _cache_timestamps.pop(f"coef:{contract}", None)
        _cache_timestamps.pop(f"cap:{contract}", None)
        logger.info("Кеш цен для контракта %s сброшен", contract)


def get_quarter_period(d: date) -> str:
    """
    Возвращает строку периода для квартала.

    Формат: "01.01.2026-31.03.2026"
    """
    year = d.year
    month = d.month

    if month <= 3:
        return f"01.01.{year}-31.03.{year}"
    elif month <= 6:
        return f"01.04.{year}-30.06.{year}"
    elif month <= 9:
        return f"01.07.{year}-30.09.{year}"
    else:
        return f"01.10.{year}-31.12.{year}"


def _load_all_coefficients(contract: str = "222") -> dict[str, dict[str, Decimal]]:
    """
    Загружает ВСЕ коэффициенты для контракта одним запросом.

    Returns:
        Словарь {period: {route: coefficient}}
    """
    pool = _get_connection_pool()
    if pool is None:
        _coefficients_loaded.add(contract)
        return {}

    conn = None
    try:
        conn = pool.getconn()
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(
                """
                SELECT route, coefficient, coef_period
                FROM routes_catalog.coefficients
                WHERE govert_contract = %s
                """,
                (contract,)
            )
            result: dict[str, dict[str, Decimal]] = {}
            for row in cur.fetchall():
                period = row['coef_period']
                if period not in result:
                    result[period] = {}
                result[period][row['route']] = Decimal(str(row['coefficient']))
            return result
    except Exception as e:
        logger.error("Ошибка загрузки коэффициентов: %s", e)
        return {}
    finally:
        if conn and pool:
            pool.putconn(conn)


def get_coefficients(period: str, contract: str = "222") -> dict[str, Decimal]:
    """
    Получает коэффициенты для всех маршрутов на период (с кэшированием).
    """
    cache_key = (period, contract)

    # Проверяем кэш
    if cache_key in _coefficients_cache:
        return _coefficients_cache[cache_key]

    # Загружаем все коэффициенты для контракта если ещё не загружены или TTL истёк
    ttl_key = f"coef:{contract}"
    if contract not in _coefficients_loaded or not _is_cache_valid(ttl_key):
        all_coefs = _load_all_coefficients(contract)
        for p, coefs in all_coefs.items():
            _coefficients_cache[(p, contract)] = coefs
        _coefficients_loaded.add(contract)
        _cache_timestamps[ttl_key] = time.monotonic()

    return _coefficients_cache.get(cache_key, {})


def _parse_period_dates(period_str: str) -> Optional[tuple[date, date]]:
    """
    Парсит строку периода в пару дат.
    Формат: 'DD.MM.YYYY<sep>DD.MM.YYYY', разделитель может быть любым
    (дефис '-', тире '–'/'—', пробел и т.д.).
    Берём первые и последние 10 символов, которые всегда являются датами.
    """
    s = period_str.strip()
    if len(s) < 21:
        return None
    try:
        a = s[:10].split('.')
        b = s[-10:].split('.')
        d_from = date(int(a[2]), int(a[1]), int(a[0]))
        d_to = date(int(b[2]), int(b[1]), int(b[0]))
        return d_from, d_to
    except (ValueError, IndexError):
        return None


def _collect_matching_periods(
    d: date,
    contract: str,
) -> list[tuple[int, dict[str, Decimal]]]:
    """Возвращает все кэшированные периоды, покрывающие дату d."""
    matching: list[tuple[int, dict[str, Decimal]]] = []
    for (period, c), coef_data in _coefficients_cache.items():
        if c != contract or not coef_data:
            continue
        parsed = _parse_period_dates(period)
        if parsed and parsed[0] <= d <= parsed[1]:
            duration = (parsed[1] - parsed[0]).days
            matching.append((duration, coef_data))
    return matching


def get_coefficients_for_date(d: date, contract: str = "222") -> dict[str, Decimal]:
    """
    Возвращает объединённые коэффициенты для всех маршрутов на дату d.

    Разные маршруты могут иметь разные периоды для одной даты
    (например, один маршрут — "01.07.2028-14.07.2028", другой — "01.07.2028-30.09.2028").
    Собираем коэффициенты из ВСЕХ периодов, покрывающих дату.
    При конфликте побеждает более короткий (специфичный) период.

    Если для d ничего не найдено — пробуем первый день того же месяца
    (покрывает случай, когда d_to = конец месяца, а коэффициенты заканчиваются раньше,
    например период контракта 01.07.2028-14.07.2028, но этап записан как 01.07-31.07).
    """
    ttl_key = f"coef:{contract}"
    if contract not in _coefficients_loaded or not _is_cache_valid(ttl_key):
        all_coefs = _load_all_coefficients(contract)
        for p, coefs in all_coefs.items():
            _coefficients_cache[(p, contract)] = coefs
        _coefficients_loaded.add(contract)
        _cache_timestamps[ttl_key] = time.monotonic()

    matching = _collect_matching_periods(d, contract)

    # Fallback: начало того же месяца (если d_to выходит за пределы последнего периода)
    if not matching:
        month_start = date(d.year, d.month, 1)
        if month_start != d:
            matching = _collect_matching_periods(month_start, contract)

    if not matching:
        return {}

    # Объединяем: сначала длинные (общий случай), потом короткие (перезаписывают)
    matching.sort(key=lambda x: x[0], reverse=True)
    merged: dict[str, Decimal] = {}
    for _, coef_data in matching:
        merged.update(coef_data)

    return merged


def get_capacities(contract: str = "222") -> dict[str, int]:
    """
    Получает вместимости для всех маршрутов (с кэшированием).
    """
    # Проверяем кэш с учётом TTL
    cache_key = f"cap:{contract}"
    if contract in _capacities_cache and _is_cache_valid(cache_key):
        return _capacities_cache[contract]

    pool = _get_connection_pool()
    if pool is None:
        _capacities_cache[contract] = {}
        return {}

    conn = None
    try:
        conn = pool.getconn()
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(
                """
                SELECT route, capacity
                FROM routes_catalog.routes
                WHERE govert_contract = %s
                """,
                (contract,)
            )
            result = {
                row['route']: int(row['capacity'])
                for row in cur.fetchall()
            }
            _capacities_cache[contract] = result
            _cache_timestamps[cache_key] = time.monotonic()
            return result
    except Exception as e:
        logger.error("Ошибка получения вместимостей: %s", e)
        _capacities_cache[contract] = {}
        return {}
    finally:
        if conn and pool:
            pool.putconn(conn)


def preload_price_data(contract: str = "222") -> None:
    """
    Предзагружает все данные для расчёта цены в кэш.
    Вызвать один раз перед массовым расчётом этапов.
    """
    # Загружаем вместимости
    get_capacities(contract)

    # Загружаем все коэффициенты
    if contract not in _coefficients_loaded:
        all_coefs = _load_all_coefficients(contract)
        for p, coefs in all_coefs.items():
            _coefficients_cache[(p, contract)] = coefs
        _coefficients_loaded.add(contract)


def calculate_route_price(
    route: str,
    km: float,
    coefficient: Decimal,
    capacity: int
) -> Decimal:
    """
    Рассчитывает цену для одного маршрута.

    Формула: км * вместимость * коэффициент
    """
    return Decimal(str(km)) * Decimal(capacity) * coefficient


def calculate_stage_price(
    route_km: dict[str, float],
    period_end_date: date,
    contract: str = "222"
) -> Optional[float]:
    """
    Рассчитывает суммарную цену этапа.

    Args:
        route_km: Словарь {route: km} - км по каждому маршруту
        period_end_date: Дата окончания периода (для определения квартала)
        contract: Номер госконтракта

    Returns:
        Суммарная цена или None если расчет невозможен
    """
    if not settings.external_db_url:
        return None

    coefficients = get_coefficients_for_date(period_end_date, contract)
    capacities = get_capacities(contract)

    if not coefficients or not capacities:
        return None

    total_price = Decimal("0")

    for route, km in route_km.items():
        # Нормализуем номер маршрута (убираем точку в конце)
        normalized_route = route.rstrip('.')

        coef = coefficients.get(normalized_route)
        cap = capacities.get(normalized_route)

        if coef is None or cap is None:
            # Пропускаем маршрут без данных
            continue

        route_price = calculate_route_price(normalized_route, km, coef, cap)
        total_price += route_price

    return float(total_price)
