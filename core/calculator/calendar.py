"""
Модуль работы с календарём.

Определяет тип дня для расчёта рейсов с учётом:
- Базового календаря (выходные, праздники)
- Переопределений для конкретных маршрутов
"""
from datetime import date, timedelta
from sqlalchemy.orm import Session

from db.models import CalendarBase, CalendarRouteOverride
from core.constants import RECURRING_HOLIDAYS


def get_weekday(d: date) -> int:
    """Возвращает день недели (1=пн, 7=вс)."""
    return d.isoweekday()


def is_recurring_holiday(d: date) -> bool:
    """Проверяет, является ли дата повторяющимся праздником."""
    return (d.month, d.day) in RECURRING_HOLIDAYS


def get_day_type(
    session: Session,
    d: date,
    contract_id: int,
    route: str
) -> int:
    """
    Определяет тип дня для расчёта рейсов.

    Приоритет:
    1. Переопределение для конкретного маршрута (CalendarRouteOverride)
    2. Базовый календарь (CalendarBase)
    3. Обычный день недели

    Args:
        session: Сессия БД
        d: Дата
        contract_id: ID контракта
        route: Номер маршрута

    Returns:
        Тип дня (1-7, где 7=воскресенье/праздник)
    """
    # 1. Проверяем переопределение для маршрута
    override = session.query(CalendarRouteOverride).filter(
        CalendarRouteOverride.date == d,
        CalendarRouteOverride.contract_id == contract_id,
        CalendarRouteOverride.route == route,
    ).first()

    if override:
        return override.treat_as

    # 2. Проверяем базовый календарь
    base = session.query(CalendarBase).filter(CalendarBase.date == d).first()

    if base and base.treat_as is not None:
        return base.treat_as

    # 3. Проверяем повторяющиеся праздники
    if is_recurring_holiday(d):
        return 7  # Считать как воскресенье

    # 4. Обычный день недели
    return get_weekday(d)


def get_day_type_batch(
    session: Session,
    dates: list[date],
    contract_id: int,
    route: str
) -> dict[date, int]:
    """
    Определяет типы дней для списка дат (оптимизированный запрос).

    Args:
        session: Сессия БД
        dates: Список дат
        contract_id: ID контракта
        route: Номер маршрута

    Returns:
        Словарь {дата: тип_дня}
    """
    if not dates:
        return {}

    min_date = min(dates)
    max_date = max(dates)

    # Загружаем переопределения одним запросом
    overrides = session.query(CalendarRouteOverride).filter(
        CalendarRouteOverride.date >= min_date,
        CalendarRouteOverride.date <= max_date,
        CalendarRouteOverride.contract_id == contract_id,
        CalendarRouteOverride.route == route,
    ).all()

    override_map = {o.date: o.treat_as for o in overrides}

    # Загружаем базовый календарь
    base_calendar = session.query(CalendarBase).filter(
        CalendarBase.date >= min_date,
        CalendarBase.date <= max_date,
    ).all()

    base_map = {b.date: b.treat_as for b in base_calendar if b.treat_as is not None}

    # Формируем результат
    result = {}
    for d in dates:
        if d in override_map:
            result[d] = override_map[d]
        elif d in base_map:
            result[d] = base_map[d]
        elif is_recurring_holiday(d):
            result[d] = 7
        else:
            result[d] = get_weekday(d)

    return result


def generate_base_calendar(
    session: Session,
    year_from: int,
    year_to: int
):
    """
    Генерирует базовый календарь на указанный период.

    Заполняет только праздничные дни из RECURRING_HOLIDAYS.

    Args:
        session: Сессия БД
        year_from: Начальный год
        year_to: Конечный год
    """
    for year in range(year_from, year_to + 1):
        for month, day in RECURRING_HOLIDAYS:
            try:
                d = date(year, month, day)
            except ValueError:
                continue

            existing = session.query(CalendarBase).filter(CalendarBase.date == d).first()
            if not existing:
                entry = CalendarBase(
                    date=d,
                    weekday=get_weekday(d),
                    is_holiday=True,
                    treat_as=7,
                    note=_get_holiday_name(month, day),
                )
                session.add(entry)


def _get_holiday_name(month: int, day: int) -> str:
    """Возвращает название праздника."""
    holidays = {
        (1, 1): "Новый год",
        (1, 2): "Новогодние каникулы",
        (1, 3): "Новогодние каникулы",
        (1, 4): "Новогодние каникулы",
        (1, 5): "Новогодние каникулы",
        (1, 6): "Новогодние каникулы",
        (1, 7): "Рождество",
        (1, 8): "Новогодние каникулы",
        (2, 23): "День защитника Отечества",
        (3, 8): "Международный женский день",
        (5, 1): "Праздник Весны и Труда",
        (5, 9): "День Победы",
        (6, 12): "День России",
        (11, 4): "День народного единства",
    }
    return holidays.get((month, day), "Праздник")


def get_dates_in_range(date_from: date, date_to: date) -> list[date]:
    """Генерирует список дат в диапазоне."""
    dates = []
    current = date_from
    while current <= date_to:
        dates.append(current)
        current += timedelta(days=1)
    return dates
