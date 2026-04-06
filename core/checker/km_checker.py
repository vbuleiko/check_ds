"""
Проверка км из Excel против рассчитанных значений.

Сравнивает данные из Приложения №12/13 с рассчитанными км.
Использует ту же логику расчёта, что и экспорт на вкладке Расчёты:
- Нормализация маршрутов (315. → 315, суммирование вариантов)
- Полный месяц (1 - последний день) через calendar.monthrange
"""
import calendar
from dataclasses import dataclass, field
from datetime import date
from typing import Optional

from sqlalchemy.orm import Session

from core.parser.km_excel import KmData, MonthlyKm
from core.calculator.kilometers import calculate_route_period
from db.models import Contract, RouteParams


# Минимальная дата для проверки (с марта 2026)
MIN_CHECK_DATE = date(2026, 3, 1)

# Допустимая погрешность (км)
KM_TOLERANCE = 0.1


@dataclass
class KmDiscrepancy:
    """Расхождение по км."""
    route: str
    year: int
    month: int
    expected_km: float  # Рассчитанные км
    actual_km: float    # Км из Excel
    diff: float         # Разница

    def __str__(self) -> str:
        month_names = [
            "", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
            "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
        ]
        month_name = month_names[self.month] if 1 <= self.month <= 12 else str(self.month)
        sign = "+" if self.diff > 0 else ""
        return (
            f"Маршрут {self.route}, {month_name} {self.year}: "
            f"рассчитано {self.expected_km:.2f}, в Excel {self.actual_km:.2f}, "
            f"разница {sign}{self.diff:.2f} км"
        )


@dataclass
class TotalKmDiscrepancy:
    """Расхождение итоговых км между Excel и JSON."""
    field_name: str  # probeg_sravnenie / probeg_etapy / probeg_etapy_avans
    excel_total: float
    json_value: float
    diff: float

    def __str__(self) -> str:
        labels = {
            "probeg_sravnenie": "Пробег (сравнение)",
            "probeg_etapy": "Пробег (этапы)",
            "probeg_etapy_avans": "Пробег (этапы с авансом)",
        }
        label = labels.get(self.field_name, self.field_name)
        sign = "+" if self.diff > 0 else ""
        return (
            f"{label}: в Excel {self.excel_total:.2f}, "
            f"в JSON {self.json_value:.2f}, "
            f"разница {sign}{self.diff:.2f} км"
        )


@dataclass
class KmCheckResult:
    """Результат проверки км."""
    discrepancies: list[KmDiscrepancy] = field(default_factory=list)
    total_discrepancies: list[TotalKmDiscrepancy] = field(default_factory=list)
    checked_periods: int = 0
    checked_routes: int = 0
    skipped_periods: int = 0  # Периоды до марта 2026
    skipped_routes: int = 0  # Маршруты, не найденные в БД

    @property
    def has_errors(self) -> bool:
        return len(self.discrepancies) > 0 or len(self.total_discrepancies) > 0

    def to_errors(self) -> list[str]:
        """Преобразует в список строк ошибок."""
        errors = [str(d) for d in self.total_discrepancies]
        errors.extend(str(d) for d in self.discrepancies)
        return errors


def normalize_route_name(route: str) -> str:
    """
    Нормализует номер маршрута:
    - Убирает точку в конце (315. -> 315)
    - Убирает суффиксы из 2+ кириллических букв (555сез -> 555, 10нов -> 10)
    - Одиночный кириллический/латинский символ сохраняется (10А -> 10А)
    """
    import re
    route = route.rstrip('.').strip()
    route = re.sub(r'[а-яёА-ЯЁ]{2,}$', '', route)
    return route


def _build_route_mapping(session: Session, contract_id: int) -> dict[str, list[str]]:
    """
    Строит маппинг: нормализованный маршрут → список оригинальных маршрутов в БД.

    Аналогично логике в export_excel.py для корректного суммирования вариантов
    маршрутов (например, "315" и "315." считаются одним маршрутом).
    """
    route_records = session.query(RouteParams.route).filter(
        RouteParams.contract_id == contract_id
    ).distinct().all()

    mapping: dict[str, list[str]] = {}
    for (route,) in route_records:
        norm = normalize_route_name(route)
        if norm not in mapping:
            mapping[norm] = []
        if route not in mapping[norm]:
            mapping[norm].append(route)

    return mapping


def _calculate_route_km(
    session: Session,
    contract_id: int,
    route_variants: list[str],
    date_from: date,
    date_to: date,
) -> float:
    """
    Рассчитывает км для маршрута, суммируя все варианты написания.

    Аналогично логике в export_excel.py: для каждого варианта маршрута
    (например, "315" и "315.") считаем отдельно и суммируем.
    """
    total_km = 0.0
    for route in route_variants:
        calc = calculate_route_period(session, contract_id, route, date_from, date_to)
        total_km += round(calc.total_km, 2)
    return total_km


def check_km_data(
    session: Session,
    contract_id: int,
    km_data: KmData,
    tolerance: float = KM_TOLERANCE,
) -> KmCheckResult:
    """
    Проверяет км из Excel против рассчитанных.

    Использует ту же логику, что и экспорт на вкладке Расчёты:
    - Нормализация маршрутов и суммирование вариантов
    - Период = полный месяц (1 - последний день по calendar.monthrange)
    """
    result = KmCheckResult()

    # Строим маппинг маршрутов из БД (как в export_excel.py)
    route_mapping = _build_route_mapping(session, contract_id)

    for monthly in km_data.monthly:
        # Пропускаем квартальные периоды (проверяем только помесячные)
        if monthly.month_end is not None:
            result.skipped_periods += 1
            continue

        # Пропускаем периоды до марта 2026
        period_date = date(monthly.year, monthly.month, 1)
        if period_date < MIN_CHECK_DATE:
            result.skipped_periods += 1
            continue

        result.checked_periods += 1

        # Используем полный месяц (как в export_excel.py)
        _, last_day = calendar.monthrange(monthly.year, monthly.month)
        date_from = date(monthly.year, monthly.month, 1)
        date_to = date(monthly.year, monthly.month, last_day)

        # Проверяем каждый маршрут
        for route, excel_km in monthly.routes.items():
            # Нормализуем маршрут из Excel
            norm_route = normalize_route_name(route)

            # Ищем варианты маршрута в БД
            route_variants = route_mapping.get(norm_route)
            if not route_variants:
                # Маршрут из Excel не найден в БД — пропускаем
                result.skipped_routes += 1
                continue

            result.checked_routes += 1

            # Рассчитываем км (суммируя все варианты, как в export_excel.py)
            calculated_km = _calculate_route_km(
                session, contract_id, route_variants, date_from, date_to
            )

            # Сравниваем
            diff = excel_km - calculated_km
            if abs(diff) > tolerance:
                result.discrepancies.append(KmDiscrepancy(
                    route=route,
                    year=monthly.year,
                    month=monthly.month,
                    expected_km=calculated_km,
                    actual_km=excel_km,
                    diff=diff,
                ))

    return result


def _check_total_km(
    km_data: KmData,
    json_data: dict,
    tolerance: float = KM_TOLERANCE,
) -> list[TotalKmDiscrepancy]:
    """
    Проверяет итоговую сумму км из Excel против probeg_* полей JSON.
    """
    if km_data.grand_total is None:
        return []

    general = json_data.get("general", {})
    discrepancies = []

    for field_name in ("probeg_sravnenie", "probeg_etapy", "probeg_etapy_avans"):
        json_value = general.get(field_name)
        if json_value is None:
            continue

        diff = km_data.grand_total - json_value
        if abs(diff) > tolerance:
            discrepancies.append(TotalKmDiscrepancy(
                field_name=field_name,
                excel_total=km_data.grand_total,
                json_value=json_value,
                diff=diff,
            ))

    return discrepancies


def check_km_for_agreement(
    session: Session,
    contract_id: int,
    json_data: dict,
) -> Optional[KmCheckResult]:
    """
    Проверяет км для ДС, если есть данные km_data.

    Args:
        session: Сессия БД
        contract_id: ID контракта
        json_data: JSON данные ДС (содержит km_data)

    Returns:
        KmCheckResult или None, если нет данных км
    """
    km_data_dict = json_data.get("km_data")
    if not km_data_dict:
        return None

    km_data = KmData.from_dict(km_data_dict)
    if not km_data.monthly:
        return None

    result = check_km_data(session, contract_id, km_data)

    # Проверяем итоговые суммы
    result.total_discrepancies = _check_total_km(km_data, json_data)

    return result
