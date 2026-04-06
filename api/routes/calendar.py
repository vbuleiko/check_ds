"""API для работы с календарём."""
from datetime import date, datetime
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session

from db.database import get_db
from db.models import CalendarBase, CalendarRouteOverride, Contract, RouteParams, RouteSeasonConfig, RouteSeasonPeriod, SeasonType
from core.calculator.calendar import generate_base_calendar, get_day_type

router = APIRouter()


@router.get("/")
async def get_calendar(
    year: int = Query(default=None),
    month: int = Query(default=None),
    contract_number: str = Query(default=None),
    db: Session = Depends(get_db)
):
    """
    Получить календарь.

    Args:
        year: Год
        month: Месяц
        contract_number: Номер контракта для фильтрации override
    """
    # Определяем период
    if year is None:
        year = datetime.now().year
    if month is None:
        # Возвращаем весь год
        start = date(year, 1, 1)
        end = date(year, 12, 31)
    else:
        start = date(year, month, 1)
        if month == 12:
            end = date(year, 12, 31)
        else:
            end = date(year, month + 1, 1)
            from datetime import timedelta
            end = end - timedelta(days=1)

    # Базовый календарь
    base_entries = db.query(CalendarBase).filter(
        CalendarBase.date >= start,
        CalendarBase.date <= end,
    ).all()

    base_map = {
        e.date: {
            "date": str(e.date),
            "weekday": e.weekday,
            "is_holiday": e.is_holiday,
            "treat_as": e.treat_as,
            "note": e.note,
        }
        for e in base_entries
    }

    # Override по всем контрактам (или по конкретному, если указан)
    overrides = []
    override_query = db.query(CalendarRouteOverride, Contract.number).join(
        Contract, CalendarRouteOverride.contract_id == Contract.id
    ).filter(
        CalendarRouteOverride.date >= start,
        CalendarRouteOverride.date <= end,
    )
    if contract_number:
        override_query = override_query.filter(Contract.number == contract_number)

    for e, c_number in override_query.all():
        overrides.append({
            "date": str(e.date),
            "route": e.route,
            "treat_as": e.treat_as,
            "source_text": e.source_text,
            "source_agreement_id": e.source_agreement_id,
            "contract_number": c_number,
        })

    # Генерируем все дни периода
    days = []
    current = start
    while current <= end:
        day_info = base_map.get(current, {
            "date": str(current),
            "weekday": current.isoweekday(),
            "is_holiday": False,
            "treat_as": None,
            "note": None,
        })
        days.append(day_info)
        from datetime import timedelta
        current = current + timedelta(days=1)

    return {
        "year": year,
        "month": month,
        "days": days,
        "overrides": overrides,
    }


@router.post("/base/{date_str}")
async def set_base_calendar(
    date_str: str,
    treat_as: int = None,
    is_holiday: bool = False,
    note: str = None,
    db: Session = Depends(get_db)
):
    """
    Установить/обновить запись базового календаря.

    Args:
        date_str: Дата в формате YYYY-MM-DD
        treat_as: Считать как день недели (1-7)
        is_holiday: Праздничный день
        note: Примечание
    """
    try:
        d = date.fromisoformat(date_str)
    except ValueError:
        raise HTTPException(400, "Неверный формат даты")

    entry = db.query(CalendarBase).filter(CalendarBase.date == d).first()

    if entry:
        entry.treat_as = treat_as
        entry.is_holiday = is_holiday
        entry.note = note
    else:
        entry = CalendarBase(
            date=d,
            weekday=d.isoweekday(),
            is_holiday=is_holiday,
            treat_as=treat_as,
            note=note,
        )
        db.add(entry)

    db.commit()
    return {"success": True}


@router.post("/override")
async def add_override(
    date_str: str,
    contract_number: str,
    route: str,
    treat_as: int,
    source_text: str = None,
    db: Session = Depends(get_db)
):
    """
    Добавить переопределение для маршрута.
    """
    try:
        d = date.fromisoformat(date_str)
    except ValueError:
        raise HTTPException(400, "Неверный формат даты")

    contract = db.query(Contract).filter(Contract.number == contract_number).first()
    if not contract:
        raise HTTPException(404, "Контракт не найден")

    # Проверяем существование
    existing = db.query(CalendarRouteOverride).filter(
        CalendarRouteOverride.date == d,
        CalendarRouteOverride.contract_id == contract.id,
        CalendarRouteOverride.route == route,
    ).first()

    if existing:
        existing.treat_as = treat_as
        existing.source_text = source_text
    else:
        entry = CalendarRouteOverride(
            date=d,
            contract_id=contract.id,
            route=route,
            treat_as=treat_as,
            source_text=source_text,
        )
        db.add(entry)

    db.commit()
    return {"success": True}


@router.post("/override-by-route")
async def add_override_by_route(
    date_str: str,
    route: str,
    treat_as: int,
    source_text: str = None,
    db: Session = Depends(get_db)
):
    """Добавить переопределение для маршрута — контракт определяется автоматически."""
    try:
        d = date.fromisoformat(date_str)
    except ValueError:
        raise HTTPException(400, "Неверный формат даты")

    # Находим контракт по маршруту через RouteParams
    rp = db.query(RouteParams).filter(RouteParams.route == route).first()
    if not rp:
        raise HTTPException(404, f"Маршрут {route} не найден ни в одном контракте")

    contract_id = rp.contract_id

    existing = db.query(CalendarRouteOverride).filter(
        CalendarRouteOverride.date == d,
        CalendarRouteOverride.contract_id == contract_id,
        CalendarRouteOverride.route == route,
    ).first()

    if existing:
        existing.treat_as = treat_as
        existing.source_text = source_text
    else:
        db.add(CalendarRouteOverride(
            date=d, contract_id=contract_id, route=route,
            treat_as=treat_as, source_text=source_text,
        ))

    db.commit()
    return {"success": True}


@router.post("/override-by-contract")
async def add_override_by_contract(
    date_str: str,
    contract_number: str,
    treat_as: int,
    source_text: str = None,
    db: Session = Depends(get_db)
):
    """Добавить переопределение для всех маршрутов ГК."""
    try:
        d = date.fromisoformat(date_str)
    except ValueError:
        raise HTTPException(400, "Неверный формат даты")

    contract = db.query(Contract).filter(Contract.number == contract_number).first()
    if not contract:
        raise HTTPException(404, "Контракт не найден")

    # Все уникальные маршруты контракта
    routes = db.query(RouteParams.route).filter(
        RouteParams.contract_id == contract.id
    ).distinct().all()

    if not routes:
        raise HTTPException(404, f"В ГК {contract_number} нет маршрутов")

    count = 0
    for (route,) in routes:
        existing = db.query(CalendarRouteOverride).filter(
            CalendarRouteOverride.date == d,
            CalendarRouteOverride.contract_id == contract.id,
            CalendarRouteOverride.route == route,
        ).first()
        if existing:
            existing.treat_as = treat_as
            existing.source_text = source_text
        else:
            db.add(CalendarRouteOverride(
                date=d, contract_id=contract.id, route=route,
                treat_as=treat_as, source_text=source_text,
            ))
        count += 1

    db.commit()
    return {"success": True, "routes_updated": count}


@router.delete("/override")
async def delete_override(
    date_str: str,
    contract_number: str,
    route: str,
    db: Session = Depends(get_db)
):
    """Удалить переопределение."""
    try:
        d = date.fromisoformat(date_str)
    except ValueError:
        raise HTTPException(400, "Неверный формат даты")

    contract = db.query(Contract).filter(Contract.number == contract_number).first()
    if not contract:
        raise HTTPException(404, "Контракт не найден")

    entry = db.query(CalendarRouteOverride).filter(
        CalendarRouteOverride.date == d,
        CalendarRouteOverride.contract_id == contract.id,
        CalendarRouteOverride.route == route,
    ).first()

    if entry:
        db.delete(entry)
        db.commit()

    return {"success": True}


@router.get("/overrides-by-route")
async def get_overrides_by_route(
    route: str,
    contract_number: str = Query(default=None),
    db: Session = Depends(get_db)
):
    """Получить все переопределения для конкретного маршрута (все даты, все контракты или конкретный)."""
    query = db.query(CalendarRouteOverride, Contract.number).join(
        Contract, CalendarRouteOverride.contract_id == Contract.id
    ).filter(
        CalendarRouteOverride.route == route,
    )
    if contract_number:
        query = query.filter(Contract.number == contract_number)

    query = query.order_by(CalendarRouteOverride.date)
    results = query.all()

    return {
        "route": route,
        "count": len(results),
        "overrides": [
            {
                "date": str(e.date),
                "treat_as": e.treat_as,
                "source_text": e.source_text,
                "source_agreement_id": e.source_agreement_id,
                "contract_number": c_number,
            }
            for e, c_number in results
        ],
    }


@router.post("/generate")
async def generate_calendar(
    year_from: int = 2022,
    year_to: int = 2028,
    db: Session = Depends(get_db)
):
    """Генерирует базовый календарь на указанный период."""
    generate_base_calendar(db, year_from, year_to)
    db.commit()
    return {"success": True, "message": f"Календарь сгенерирован на {year_from}-{year_to}"}


# =============================================================================
# Сезонные конфигурации маршрутов
# =============================================================================

def _season_config_to_dict(c: RouteSeasonConfig) -> dict:
    def fmt(m, d):
        return f"{d:02d}.{m:02d}"
    return {
        "route": c.route,
        "winter_start": fmt(c.winter_start_month, c.winter_start_day),
        "winter_end": fmt(c.winter_end_month, c.winter_end_day),
        "summer_start": fmt(c.summer_start_month, c.summer_start_day),
        "summer_end": fmt(c.summer_end_month, c.summer_end_day),
    }


def _parse_day_month(s: str) -> tuple[int, int]:
    """Парсит строку вида 'DD.MM' → (month, day)."""
    parts = s.strip().split(".")
    if len(parts) != 2:
        raise ValueError(f"Неверный формат даты: {s}")
    day, month = int(parts[0]), int(parts[1])
    if not (1 <= month <= 12) or not (1 <= day <= 31):
        raise ValueError(f"Недопустимые значения: {s}")
    return month, day


@router.get("/season-configs")
async def list_season_configs(db: Session = Depends(get_db)):
    """Список всех сезонных конфигураций маршрутов."""
    configs = db.query(RouteSeasonConfig).order_by(RouteSeasonConfig.route).all()
    return [_season_config_to_dict(c) for c in configs]


@router.post("/season-configs")
async def create_season_config(
    route: str,
    winter_start: str,
    winter_end: str,
    summer_start: str,
    summer_end: str,
    db: Session = Depends(get_db)
):
    """Создать новую сезонную конфигурацию маршрута."""
    existing = db.query(RouteSeasonConfig).filter(RouteSeasonConfig.route == route).first()
    if existing:
        raise HTTPException(400, f"Конфигурация для маршрута {route} уже существует")

    try:
        wsm, wsd = _parse_day_month(winter_start)
        wem, wed = _parse_day_month(winter_end)
        ssm, ssd = _parse_day_month(summer_start)
        sem, sed = _parse_day_month(summer_end)
    except ValueError as e:
        raise HTTPException(400, str(e))

    config = RouteSeasonConfig(
        route=route,
        winter_start_month=wsm, winter_start_day=wsd,
        winter_end_month=wem, winter_end_day=wed,
        summer_start_month=ssm, summer_start_day=ssd,
        summer_end_month=sem, summer_end_day=sed,
    )
    db.add(config)
    db.commit()
    return _season_config_to_dict(config)


@router.put("/season-configs/{route}")
async def update_season_config(
    route: str,
    winter_start: str,
    winter_end: str,
    summer_start: str,
    summer_end: str,
    db: Session = Depends(get_db)
):
    """Обновить сезонную конфигурацию маршрута."""
    config = db.query(RouteSeasonConfig).filter(RouteSeasonConfig.route == route).first()
    if not config:
        raise HTTPException(404, f"Конфигурация для маршрута {route} не найдена")

    try:
        wsm, wsd = _parse_day_month(winter_start)
        wem, wed = _parse_day_month(winter_end)
        ssm, ssd = _parse_day_month(summer_start)
        sem, sed = _parse_day_month(summer_end)
    except ValueError as e:
        raise HTTPException(400, str(e))

    config.winter_start_month = wsm
    config.winter_start_day = wsd
    config.winter_end_month = wem
    config.winter_end_day = wed
    config.summer_start_month = ssm
    config.summer_start_day = ssd
    config.summer_end_month = sem
    config.summer_end_day = sed
    db.commit()
    return _season_config_to_dict(config)


@router.delete("/season-configs/{route}")
async def delete_season_config(route: str, db: Session = Depends(get_db)):
    """Удалить сезонную конфигурацию маршрута."""
    config = db.query(RouteSeasonConfig).filter(RouteSeasonConfig.route == route).first()
    if not config:
        raise HTTPException(404, f"Конфигурация для маршрута {route} не найдена")
    db.delete(config)
    db.commit()
    return {"success": True}


# =============================================================================
# Конкретные периоды по годам (RouteSeasonPeriod)
# =============================================================================

def _period_to_dict(p: RouteSeasonPeriod) -> dict:
    return {
        "id": p.id,
        "route": p.route,
        "season": p.season.value,
        "date_from": str(p.date_from),
        "date_to": str(p.date_to),
    }


@router.get("/season-periods")
async def list_season_periods(
    route: Optional[str] = Query(default=None),
    db: Session = Depends(get_db)
):
    """Список конкретных сезонных периодов. Опционально фильтр по маршруту."""
    q = db.query(RouteSeasonPeriod)
    if route:
        q = q.filter(RouteSeasonPeriod.route == route)
    periods = q.order_by(RouteSeasonPeriod.route, RouteSeasonPeriod.date_from).all()
    return [_period_to_dict(p) for p in periods]


@router.post("/season-periods")
async def create_season_period(
    route: str,
    season: str,
    date_from: str,
    date_to: str,
    db: Session = Depends(get_db)
):
    """Создать конкретный сезонный период."""
    try:
        season_type = SeasonType(season)
    except ValueError:
        raise HTTPException(400, f"Неверный сезон: {season}. Допустимые: winter, summer")
    try:
        df = date.fromisoformat(date_from)
        dt = date.fromisoformat(date_to)
    except ValueError:
        raise HTTPException(400, "Неверный формат даты, ожидается YYYY-MM-DD")
    if dt <= df:
        raise HTTPException(400, "date_to должна быть позже date_from")

    existing = db.query(RouteSeasonPeriod).filter(
        RouteSeasonPeriod.route == route,
        RouteSeasonPeriod.season == season_type,
        RouteSeasonPeriod.date_from == df,
    ).first()
    if existing:
        raise HTTPException(400, "Период с такими параметрами уже существует")

    period = RouteSeasonPeriod(route=route, season=season_type, date_from=df, date_to=dt)
    db.add(period)
    db.commit()
    return _period_to_dict(period)


@router.put("/season-periods/{period_id}")
async def update_season_period(
    period_id: int,
    season: str,
    date_from: str,
    date_to: str,
    db: Session = Depends(get_db)
):
    """Обновить конкретный сезонный период."""
    period = db.query(RouteSeasonPeriod).filter(RouteSeasonPeriod.id == period_id).first()
    if not period:
        raise HTTPException(404, "Период не найден")
    try:
        season_type = SeasonType(season)
    except ValueError:
        raise HTTPException(400, f"Неверный сезон: {season}")
    try:
        df = date.fromisoformat(date_from)
        dt = date.fromisoformat(date_to)
    except ValueError:
        raise HTTPException(400, "Неверный формат даты")
    if dt <= df:
        raise HTTPException(400, "date_to должна быть позже date_from")

    period.season = season_type
    period.date_from = df
    period.date_to = dt
    db.commit()
    return _period_to_dict(period)


@router.delete("/season-periods/{period_id}")
async def delete_season_period(period_id: int, db: Session = Depends(get_db)):
    """Удалить конкретный сезонный период."""
    period = db.query(RouteSeasonPeriod).filter(RouteSeasonPeriod.id == period_id).first()
    if not period:
        raise HTTPException(404, "Период не найден")
    db.delete(period)
    db.commit()
    return {"success": True}


@router.post("/season-periods/generate")
async def generate_season_periods(
    route: str,
    year_from: int,
    year_to: int,
    db: Session = Depends(get_db)
):
    """Сгенерировать конкретные периоды для маршрута из шаблона (RouteSeasonConfig)."""
    from core.constants import ROUTE_SEASON_PERIODS

    config = db.query(RouteSeasonConfig).filter(RouteSeasonConfig.route == route).first()
    if config:
        ws_m, ws_d = config.winter_start_month, config.winter_start_day
        we_m, we_d = config.winter_end_month, config.winter_end_day
        ss_m, ss_d = config.summer_start_month, config.summer_start_day
        se_m, se_d = config.summer_end_month, config.summer_end_day
    elif route.upper() in ROUTE_SEASON_PERIODS:
        rsp = ROUTE_SEASON_PERIODS[route.upper()]
        ws_m, ws_d, we_m, we_d = rsp["winter"]
        ss_m, ss_d, se_m, se_d = rsp["summer"]
    else:
        raise HTTPException(404, f"Шаблон для маршрута {route} не найден")

    winter_cross = ws_m > we_m
    summer_cross = ss_m > se_m
    created = 0
    for year in range(year_from, year_to + 1):
        for season_type, sfm, sfd, etm, etd, cross in [
            (SeasonType.WINTER, ws_m, ws_d, we_m, we_d, winter_cross),
            (SeasonType.SUMMER, ss_m, ss_d, se_m, se_d, summer_cross),
        ]:
            df = date(year, sfm, sfd)
            dt = date(year + 1 if cross else year, etm, etd)
            existing = db.query(RouteSeasonPeriod).filter(
                RouteSeasonPeriod.route == route,
                RouteSeasonPeriod.season == season_type,
                RouteSeasonPeriod.date_from == df,
            ).first()
            if not existing:
                db.add(RouteSeasonPeriod(route=route, season=season_type, date_from=df, date_to=dt))
                created += 1
    db.commit()
    return {"success": True, "created": created}
