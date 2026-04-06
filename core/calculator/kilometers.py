"""
Модуль расчёта километров и рейсов.

Основной расчёт транспортной работы с учётом:
- Параметров маршрута (рейсы по типам дней, протяжённость)
- Календаря (праздники, переопределения)
- Сезонности (зима/лето)
"""
from datetime import date
from dataclasses import dataclass, field
from typing import Optional

from sqlalchemy import case
from sqlalchemy.orm import Session

from db.models import RouteParams, RouteTrips, Contract, SeasonType, RouteEndDate, CalendarRouteOverride
from core.calculator.calendar import get_day_type_batch, get_dates_in_range
from core.constants import SEASON_PERIODS, ROUTE_SEASON_PERIODS


@dataclass
class DayCalculation:
    """Расчёт за один день."""
    date: date
    day_type: int  # 1-7
    forward_trips: int = 0
    reverse_trips: int = 0
    forward_km: float = 0.0
    reverse_km: float = 0.0
    route_params_id: Optional[int] = None  # ID записи route_params

    @property
    def total_trips(self) -> int:
        return self.forward_trips + self.reverse_trips

    @property
    def total_km(self) -> float:
        return self.forward_km + self.reverse_km


@dataclass
class ParamsSegment:
    """Сегмент расчёта по одной записи route_params."""
    route_params_id: Optional[int]
    date_from: date
    date_to: date
    source_appendix: Optional[str] = None
    length_forward: float = 0.0
    length_reverse: float = 0.0
    days_count: int = 0
    forward_trips: int = 0
    reverse_trips: int = 0
    forward_km: float = 0.0
    reverse_km: float = 0.0

    @property
    def total_km(self) -> float:
        return self.forward_km + self.reverse_km

    def to_dict(self) -> dict:
        return {
            "route_params_id": self.route_params_id,
            "date_from": str(self.date_from),
            "date_to": str(self.date_to),
            "source": self.source_appendix,
            "length_forward": self.length_forward,
            "length_reverse": self.length_reverse,
            "days_count": self.days_count,
            "forward_trips": self.forward_trips,
            "reverse_trips": self.reverse_trips,
            "forward_km": round(self.forward_km, 2),
            "reverse_km": round(self.reverse_km, 2),
            "total_km": round(self.total_km, 2),
        }


@dataclass
class PeriodCalculation:
    """Расчёт за период."""
    route: str
    date_from: date
    date_to: date
    days: list[DayCalculation] = field(default_factory=list)

    # Итоги по типам дней
    by_day_type: dict[int, dict] = field(default_factory=dict)

    # Сегменты по разным route_params (показывает какие параметры использовались)
    segments: list[ParamsSegment] = field(default_factory=list)

    @property
    def total_forward_trips(self) -> int:
        return sum(d.forward_trips for d in self.days)

    @property
    def total_reverse_trips(self) -> int:
        return sum(d.reverse_trips for d in self.days)

    @property
    def total_trips(self) -> int:
        return self.total_forward_trips + self.total_reverse_trips

    @property
    def total_forward_km(self) -> float:
        return sum(d.forward_km for d in self.days)

    @property
    def total_reverse_km(self) -> float:
        return sum(d.reverse_km for d in self.days)

    @property
    def total_km(self) -> float:
        return self.total_forward_km + self.total_reverse_km

    def to_dict(self) -> dict:
        return {
            "route": self.route,
            "date_from": str(self.date_from),
            "date_to": str(self.date_to),
            "total_trips": self.total_trips,
            "total_forward_trips": self.total_forward_trips,
            "total_reverse_trips": self.total_reverse_trips,
            "total_km": round(self.total_km, 2),
            "total_forward_km": round(self.total_forward_km, 2),
            "total_reverse_km": round(self.total_reverse_km, 2),
            "by_day_type": self.by_day_type,
            "segments": [s.to_dict() for s in self.segments],
        }


def get_route_end_date(
    session: Session,
    contract_id: int,
    route: str
) -> Optional[date]:
    """
    Получает дату окончания работы маршрута.

    Returns:
        Дата окончания или None, если маршрут работает до конца контракта.
    """
    # Нормализуем маршрут (убираем точку)
    normalized_route = route.rstrip('.')

    route_end = session.query(RouteEndDate).filter(
        RouteEndDate.contract_id == contract_id,
        RouteEndDate.route == normalized_route,
    ).first()

    return route_end.end_date if route_end else None


def get_season_for_date(d: date, route: str = None) -> SeasonType:
    """Определяет сезон для даты. Если передан маршрут — использует его конкретные периоды."""
    if route:
        rsp = ROUTE_SEASON_PERIODS.get(route.strip().upper())
        if rsp:
            ws_m, ws_d, we_m, we_d = rsp["winter"]
            # Зима пересекает границу года (например, нояб–апр или сент–май)
            if ws_m > we_m:
                in_winter = (
                    (d.month == ws_m and d.day >= ws_d) or
                    (d.month > ws_m) or
                    (d.month < we_m) or
                    (d.month == we_m and d.day <= we_d)
                )
            else:
                # Зима в пределах одного года
                in_winter = (
                    (d.month > ws_m or (d.month == ws_m and d.day >= ws_d)) and
                    (d.month < we_m or (d.month == we_m and d.day <= we_d))
                )
            return SeasonType.WINTER if in_winter else SeasonType.SUMMER

    # Fallback: общий сезонный календарь (16.11 – 14.04 зима)
    winter = SEASON_PERIODS["winter"]
    if (d.month == winter["start_month"] and d.day >= winter["start_day"]) or \
       (d.month > winter["start_month"]) or \
       (d.month < winter["end_month"]) or \
       (d.month == winter["end_month"] and d.day <= winter["end_day"]):
        return SeasonType.WINTER

    return SeasonType.SUMMER


def get_route_params_for_date(
    session: Session,
    contract_id: int,
    route: str,
    d: date,
    season: Optional[SeasonType] = None
) -> Optional[RouteParams]:
    """
    Получает параметры маршрута на указанную дату.

    Находит последнюю запись, действующую на эту дату.
    """
    # Определяем сезон, если не указан
    if season is None:
        season = get_season_for_date(d, route)

    # Ищем параметры:
    # 1. date_from <= d
    # 2. date_to IS NULL OR date_to >= d
    # 3. season = season OR season = ALL_YEAR
    # Приоритет: сначала по date_from desc, затем сезонные записи над ALL_YEAR,
    # затем по id desc (более поздняя запись побеждает при совпадении date_from)

    season_priority = case(
        (RouteParams.season == SeasonType.ALL_YEAR, 0),
        else_=1,
    )

    query = session.query(RouteParams).filter(
        RouteParams.contract_id == contract_id,
        RouteParams.route == route,
        RouteParams.date_from <= d,
    ).filter(
        (RouteParams.date_to.is_(None)) | (RouteParams.date_to >= d)
    ).filter(
        (RouteParams.season == season) | (RouteParams.season == SeasonType.ALL_YEAR)
    ).order_by(RouteParams.date_from.desc(), season_priority.desc(), RouteParams.id.desc())

    return query.first()


def get_route_params_nearest_future(
    session: Session,
    contract_id: int,
    route: str,
    d: date,
    season: Optional[SeasonType] = None
) -> Optional[RouteParams]:
    """
    Находит ближайшие будущие параметры маршрута (date_from > d).

    Используется как fallback, когда на дату есть переопределение календаря,
    но нет действующих route_params. Это значит, что маршрут должен работать
    по графику переопределённого дня, но параметры (протяжённость, рейсы)
    ещё не вступили в силу формально.
    """
    if season is None:
        season = get_season_for_date(d, route)

    season_priority = case(
        (RouteParams.season == SeasonType.ALL_YEAR, 0),
        else_=1,
    )

    query = session.query(RouteParams).filter(
        RouteParams.contract_id == contract_id,
        RouteParams.route == route,
        RouteParams.date_from > d,
    ).filter(
        (RouteParams.season == season) | (RouteParams.season == SeasonType.ALL_YEAR)
    ).order_by(RouteParams.date_from.asc(), season_priority.desc(), RouteParams.id.desc())

    return query.first()


def get_trips_for_day_type(
    route_params: RouteParams,
    day_type: int
) -> tuple[int, int]:
    """
    Получает количество рейсов для типа дня.

    Args:
        route_params: Параметры маршрута
        day_type: Тип дня (1-7)

    Returns:
        (forward_trips, reverse_trips)
    """
    if not route_params or not route_params.trips:
        return 0, 0

    # Ищем RouteTrips, в weekdays которого есть day_type
    for trip in route_params.trips:
        if day_type in trip.weekdays:
            return trip.forward_number, trip.reverse_number

    return 0, 0


def calculate_route_period(
    session: Session,
    contract_id: int,
    route: str,
    date_from: date,
    date_to: date
) -> PeriodCalculation:
    """
    Рассчитывает километраж маршрута за период.

    Args:
        session: Сессия БД
        contract_id: ID контракта
        route: Номер маршрута
        date_from: Начало периода
        date_to: Конец периода

    Returns:
        PeriodCalculation с детализацией по дням
    """
    result = PeriodCalculation(
        route=route,
        date_from=date_from,
        date_to=date_to,
    )

    # Проверяем дату окончания маршрута
    route_end_date = get_route_end_date(session, contract_id, route)

    # Корректируем конец периода, если маршрут заканчивается раньше
    effective_date_to = date_to
    if route_end_date and route_end_date < date_to:
        effective_date_to = route_end_date

    # Если маршрут уже закончился до начала периода
    if route_end_date and route_end_date < date_from:
        return result

    # Получаем все даты периода
    dates = get_dates_in_range(date_from, effective_date_to)
    if not dates:
        return result

    # Получаем типы дней одним запросом
    day_types = get_day_type_batch(session, dates, contract_id, route)

    # Загружаем даты с переопределениями календаря (для fallback к будущим параметрам)
    override_dates = set()
    overrides = session.query(CalendarRouteOverride.date).filter(
        CalendarRouteOverride.date >= date_from,
        CalendarRouteOverride.date <= effective_date_to,
        CalendarRouteOverride.contract_id == contract_id,
        CalendarRouteOverride.route == route,
    ).all()
    override_dates = {o.date for o in overrides}

    # Кэш параметров по датам
    params_cache: dict[date, RouteParams] = {}

    # Текущий сегмент
    current_segment: Optional[ParamsSegment] = None

    for d in dates:
        day_type = day_types.get(d, d.isoweekday())

        # Получаем параметры маршрута
        if d not in params_cache:
            params = get_route_params_for_date(session, contract_id, route, d)
            # Если параметры не найдены, но есть переопределение календаря —
            # ищем ближайшие будущие параметры. Переопределение дня означает,
            # что маршрут должен работать, просто по другому графику.
            if params is None and d in override_dates:
                params = get_route_params_nearest_future(session, contract_id, route, d)
            params_cache[d] = params
        else:
            params = params_cache[d]

        params_id = params.id if params else None

        # Проверяем, нужно ли начать новый сегмент
        if current_segment is None or current_segment.route_params_id != params_id:
            # Закрываем предыдущий сегмент
            if current_segment is not None:
                result.segments.append(current_segment)

            # Создаём новый сегмент
            current_segment = ParamsSegment(
                route_params_id=params_id,
                date_from=d,
                date_to=d,
                source_appendix=params.source_appendix if params else None,
                length_forward=params.length_forward or 0 if params else 0,
                length_reverse=params.length_reverse or 0 if params else 0,
            )
        else:
            # Расширяем текущий сегмент
            current_segment.date_to = d

        if not params:
            # Нет параметров на эту дату
            result.days.append(DayCalculation(date=d, day_type=day_type, route_params_id=None))
            current_segment.days_count += 1
            continue

        # Получаем рейсы
        forward_trips, reverse_trips = get_trips_for_day_type(params, day_type)

        # Рассчитываем км
        length_forward = params.length_forward or 0
        length_reverse = params.length_reverse or 0

        forward_km = forward_trips * length_forward
        reverse_km = reverse_trips * length_reverse

        day_calc = DayCalculation(
            date=d,
            day_type=day_type,
            forward_trips=forward_trips,
            reverse_trips=reverse_trips,
            forward_km=forward_km,
            reverse_km=reverse_km,
            route_params_id=params.id,
        )
        result.days.append(day_calc)

        # Обновляем сегмент
        current_segment.days_count += 1
        current_segment.forward_trips += forward_trips
        current_segment.reverse_trips += reverse_trips
        current_segment.forward_km += forward_km
        current_segment.reverse_km += reverse_km

        # Агрегируем по типам дней
        if day_type not in result.by_day_type:
            result.by_day_type[day_type] = {
                "days_count": 0,
                "forward_trips": 0,
                "reverse_trips": 0,
                "forward_km": 0.0,
                "reverse_km": 0.0,
            }

        result.by_day_type[day_type]["days_count"] += 1
        result.by_day_type[day_type]["forward_trips"] += forward_trips
        result.by_day_type[day_type]["reverse_trips"] += reverse_trips
        result.by_day_type[day_type]["forward_km"] += forward_km
        result.by_day_type[day_type]["reverse_km"] += reverse_km

    # Добавляем последний сегмент
    if current_segment is not None:
        result.segments.append(current_segment)

    return result


def calculate_contract_period(
    session: Session,
    contract_id: int,
    date_from: date,
    date_to: date,
    routes: Optional[list[str]] = None
) -> dict[str, PeriodCalculation]:
    """
    Рассчитывает километраж по всем маршрутам контракта за период.

    Args:
        session: Сессия БД
        contract_id: ID контракта
        date_from: Начало периода
        date_to: Конец периода
        routes: Список маршрутов (если None — все маршруты контракта)

    Returns:
        Словарь {route: PeriodCalculation}
    """
    # Если маршруты не указаны, берём все из route_params
    if routes is None:
        route_records = session.query(RouteParams.route).filter(
            RouteParams.contract_id == contract_id
        ).distinct().all()
        routes = [r[0] for r in route_records]

    results = {}
    for route in routes:
        results[route] = calculate_route_period(
            session, contract_id, route, date_from, date_to
        )

    return results
