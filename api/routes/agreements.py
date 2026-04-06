"""API для работы с дополнительными соглашениями."""
import calendar as cal_module
from datetime import date, datetime, timedelta, timezone
import re
from fastapi import APIRouter, Depends, HTTPException, Request
from sqlalchemy.orm import Session

from db.database import get_db
from db.models import Agreement, AgreementStatus, Contract, RouteParams, RouteTrips, CalendarRouteOverride, SeasonType, CalculatedStage
from services.stages_calculator import recalculate_stages
from core.constants import ROUTE_SEASON_PERIODS, get_weekdays_for_type_extended, detect_mid_season_change, parse_point_to_weekday
from core.utils import parse_date
from core.parser.km_excel import KmData
from core.checker.km_checker import MIN_CHECK_DATE
from core.calculator.export_excel import get_contract_routes, normalize_route
from core.calculator.kilometers import calculate_route_period


router = APIRouter()


@router.get("/")
async def list_agreements(
    contract_number: str = None,
    status: str = None,
    ds_type: str = None,
    db: Session = Depends(get_db)
):
    """Список ДС с фильтрацией."""
    query = db.query(Agreement)

    if contract_number:
        query = query.join(Agreement.contract).filter(
            Contract.number == contract_number
        )

    if status:
        try:
            status_enum = AgreementStatus(status)
            query = query.filter(Agreement.status == status_enum)
        except ValueError:
            pass

    agreements = query.order_by(Agreement.created_at.desc()).all()

    # Фильтрация по типу ДС (в Python, т.к. тип хранится внутри JSON-поля)
    if ds_type:
        if ds_type == "vysvobozhdenie":
            agreements = [a for a in agreements if (a.json_data or {}).get("type") == "vysvobozhdenie"]
        elif ds_type == "params":
            agreements = [a for a in agreements if (a.json_data or {}).get("type") != "vysvobozhdenie"]

    def _has_seasonal_warnings(a) -> bool:
        warnings = a.check_warnings or []
        return any(
            "возможна ошибка в графиках" in w or "изменён сезонный график" in w
            for w in warnings
        )

    def _ds_type(a) -> str:
        jd = a.json_data or {}
        if jd.get("type") == "vysvobozhdenie":
            return "Высвобождение"
        return "Изменение параметров"

    def _has_embedded_vysvobozhdenie(a) -> bool:
        jd = a.json_data or {}
        return jd.get("type") != "vysvobozhdenie" and "vysvobozhdenie" in jd

    return [
        {
            "id": a.id,
            "contract_number": a.contract.number if a.contract else None,
            "ds_number": a.number,
            "ds_type": _ds_type(a),
            "status": a.status.value,
            "date_signed": str(a.date_signed) if a.date_signed else None,
            "created_at": str(a.created_at),
            "errors_count": len(a.check_errors) if a.check_errors else 0,
            "warnings_count": len(a.check_warnings) if a.check_warnings else 0,
            "has_seasonal_warnings": _has_seasonal_warnings(a),
            "has_embedded_vysvobozhdenie": _has_embedded_vysvobozhdenie(a),
        }
        for a in agreements
    ]


@router.get("/vysvobozhdenie-history")
async def get_vysvobozhdenie_history(
    contract_number: str = None,
    db: Session = Depends(get_db)
):
    """История всех высвобождений — как отдельных ДС, так и встроенных в ДС на изменение параметров."""
    query = db.query(Agreement)
    if contract_number:
        query = query.join(Agreement.contract).filter(Contract.number == contract_number)
    agreements = query.order_by(Agreement.created_at.desc()).all()

    result = []
    for a in agreements:
        jd = a.json_data or {}
        vysv = jd.get("vysvobozhdenie")
        if not vysv:
            continue

        is_standalone = jd.get("type") == "vysvobozhdenie"
        result.append({
            "agreement_id": a.id,
            "contract_number": a.contract.number if a.contract else None,
            "ds_number": a.number,
            "is_standalone": is_standalone,
            "closed_stage": vysv.get("closed_stage"),
            "closed_amount": vysv.get("closed_amount"),
            "new_contract_price": vysv.get("new_contract_price") or (
                (jd.get("general") or {}).get("sum_text") if not is_standalone else None
            ),
            "created_at": str(a.created_at),
        })

    return result


@router.get("/{agreement_id}")
async def get_agreement(agreement_id: int, db: Session = Depends(get_db)):
    """Получить ДС по ID."""
    agreement = db.query(Agreement).filter(Agreement.id == agreement_id).first()
    if not agreement:
        raise HTTPException(404, "ДС не найден")

    return {
        "id": agreement.id,
        "contract_number": agreement.contract.number if agreement.contract else None,
        "ds_number": agreement.number,
        "status": agreement.status.value,
        "date_signed": str(agreement.date_signed) if agreement.date_signed else None,
        "json_data": agreement.json_data,
        "check_errors": agreement.check_errors or [],
        "check_warnings": agreement.check_warnings or [],
        "total_km_before": agreement.total_km_before,
        "total_km_after": agreement.total_km_after,
        "calculation_details": agreement.calculation_details,
        "created_at": str(agreement.created_at),
        "updated_at": str(agreement.updated_at),
    }


@router.get("/{agreement_id}/km-check")
async def get_km_check(agreement_id: int, db: Session = Depends(get_db)):
    """Возвращает сравнение км из Excel и расчётных значений по периодам."""
    agreement = db.query(Agreement).filter(Agreement.id == agreement_id).first()
    if not agreement:
        raise HTTPException(404, "ДС не найден")

    json_data = agreement.json_data or {}
    km_data_dict = json_data.get("km_data")
    if not km_data_dict:
        return {"rows": []}

    km_data = KmData.from_dict(km_data_dict)
    contract_id = agreement.contract_id

    # Используем ту же логику нормализации маршрутов что и экспорт Excel
    if contract_id:
        norm_routes = get_contract_routes(db, contract_id, normalize=True)
        orig_routes = get_contract_routes(db, contract_id, normalize=False)
        route_mapping: dict[str, list[str]] = {}
        for orig in orig_routes:
            norm = normalize_route(orig)
            if norm not in route_mapping:
                route_mapping[norm] = []
            route_mapping[norm].append(orig)
    else:
        norm_routes = []
        route_mapping = {}

    has_route_data = bool(norm_routes)

    MONTH_NAMES = [
        "", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
    ]

    rows = []
    for m in km_data.monthly:
        period_date = date(m.year, m.month, 1)
        is_quarterly = m.month_end is not None
        before_cutoff = period_date < MIN_CHECK_DATE

        can_check = not is_quarterly and not before_cutoff and has_route_data

        skip_reason = None
        if is_quarterly:
            skip_reason = "quarterly"
        elif before_cutoff:
            skip_reason = "before_cutoff"
        elif not has_route_data:
            skip_reason = "no_route_data"

        calculated_km = None
        if can_check:
            _, last_day = cal_module.monthrange(m.year, m.month)
            date_from = date(m.year, m.month, 1)
            date_to = date(m.year, m.month, last_day)
            # Точная копия логики из export_excel.py:
            # для каждого нормализованного маршрута суммируем все оригиналы (315 + 315.)
            total_calc = 0.0
            for norm_route in norm_routes:
                for orig_route in route_mapping.get(norm_route, []):
                    calc = calculate_route_period(db, contract_id, orig_route, date_from, date_to)
                    total_calc += round(calc.total_km, 2)
            calculated_km = round(total_calc, 2)

        month_name = MONTH_NAMES[m.month] if 1 <= m.month <= 12 else str(m.month)
        rows.append({
            "period": f"{month_name} {m.year}",
            "period_days": f"{m.period_start}-{m.period_end}",
            "routes_count": len(m.routes),
            "excel_km": m.total,
            "calculated_km": calculated_km,
            "can_check": can_check,
            "skip_reason": skip_reason,
        })

    return {"rows": rows, "has_route_data": has_route_data}


@router.patch("/{agreement_id}/json")
async def update_agreement_json(
    agreement_id: int,
    request: Request,
    db: Session = Depends(get_db)
):
    """Обновить json_data соглашения (ручное редактирование)."""
    data = await request.json()

    agreement = db.query(Agreement).filter(Agreement.id == agreement_id).first()
    if not agreement:
        raise HTTPException(404, "ДС не найден")
    if agreement.status == AgreementStatus.APPLIED:
        raise HTTPException(400, "Нельзя редактировать применённый ДС")

    agreement.json_data = data
    agreement.updated_at = datetime.now(timezone.utc)
    db.commit()
    return {"success": True}


@router.post("/{agreement_id}/apply")
async def apply_agreement(agreement_id: int, db: Session = Depends(get_db)):
    """
    Применить ДС — добавить данные в route_params и calendar_override.
    """
    agreement = db.query(Agreement).filter(Agreement.id == agreement_id).first()
    if not agreement:
        raise HTTPException(404, "ДС не найден")

    if not agreement.contract:
        raise HTTPException(400, "ДС не привязан к контракту")

    if agreement.status == AgreementStatus.APPLIED:
        raise HTTPException(400, "ДС уже применён")

    json_data = agreement.json_data or {}
    ds_number = agreement.number
    contract_id = agreement.contract_id

    added_params = 0
    added_overrides = 0
    seasonal_changes: list[str] = []

    # 1. Добавляем appendices в route_params
    appendices = json_data.get("appendices", {})
    for key, app in appendices.items():
        route = app.get("route")
        if not route:
            continue

        date_from = parse_date(app.get("date_from")) or parse_date(app.get("date_on"))
        date_to = parse_date(app.get("date_to"))

        if not date_from:
            continue

        # Определяем сезон
        season = SeasonType.ALL_YEAR
        has_winter = app.get("period_winter") and app["period_winter"].get("num_of_types", 0) > 0
        has_summer = app.get("period_summer") and app["period_summer"].get("num_of_types", 0) > 0

        # Если есть сезонные периоды, создаём отдельные записи
        periods_to_add = []
        if has_winter or has_summer:
            if has_winter:
                winter_data = app["period_winter"]
                periods_to_add.append((SeasonType.WINTER, winter_data))
            if has_summer:
                summer_data = app["period_summer"]
                periods_to_add.append((SeasonType.SUMMER, summer_data))
        else:
            # Обычное приложение без сезонов
            periods_to_add.append((SeasonType.ALL_YEAR, app))

        for season_type, period_data in periods_to_add:
            # Создаём RouteParams
            route_params = RouteParams(
                contract_id=contract_id,
                route=route,
                date_from=date_from,
                date_to=date_to,
                season=season_type,
                length_forward=app.get("length_forward"),
                length_reverse=app.get("length_reverse"),
                source_agreement_id=agreement.id,
                source_appendix=f"Приложение № {app.get('appendix_num', key)} ДС{ds_number}",
            )
            db.add(route_params)
            db.flush()  # Получаем ID

            # Добавляем типы дней (рейсы)
            num_types = period_data.get("num_of_types", 0)
            for i in range(1, num_types + 1):
                type_name = period_data.get(f"type_{i}_name", "")
                forward_num = period_data.get(f"type_{i}_forward_number", 0) or 0
                reverse_num = period_data.get(f"type_{i}_reverse_number", 0) or 0

                weekdays = get_weekdays_for_type_extended(type_name)
                if weekdays and (forward_num > 0 or reverse_num > 0):
                    trip = RouteTrips(
                        route_params_id=route_params.id,
                        day_type_name=type_name,
                        weekdays=weekdays,
                        forward_number=int(forward_num),
                        reverse_number=int(reverse_num),
                    )
                    db.add(trip)

            added_params += 1

        # Определяем, является ли изменение середино-сезонным
        if (has_winter or has_summer) and date_from:
            msg = detect_mid_season_change(route, date_from)
            if msg and msg not in seasonal_changes:
                seasonal_changes.append(msg)

    # 2. Добавляем изменения без приложений в calendar_override
    for change_list in [
        json_data.get("change_with_money_no_appendix", []),
        json_data.get("change_without_money_no_appendix", []),
    ]:
        for change in change_list:
            route = change.get("route")
            date_on = parse_date(change.get("date_on"))
            date_from = parse_date(change.get("date_from"))
            date_to = parse_date(change.get("date_to"))
            point_text = change.get("point", "")

            if not route:
                continue

            # Определяем как считать день
            treat_as = parse_point_to_weekday(point_text)
            if "выходн" in point_text.lower() and not treat_as:
                treat_as = 7

            if not treat_as:
                continue

            # Собираем список дат для создания override
            dates_to_add = []
            if date_from and date_to:
                # Диапазон дат
                current = date_from
                while current <= date_to:
                    dates_to_add.append(current)
                    current += timedelta(days=1)
            elif date_on:
                # Одна конкретная дата
                dates_to_add.append(date_on)
            else:
                continue

            for override_date in dates_to_add:
                # Проверяем, нет ли уже такой записи
                existing = db.query(CalendarRouteOverride).filter(
                    CalendarRouteOverride.date == override_date,
                    CalendarRouteOverride.contract_id == contract_id,
                    CalendarRouteOverride.route == route,
                ).first()
                if existing:
                    # Обновляем существующий override — привязываем к текущему ДС
                    existing.treat_as = treat_as
                    existing.source_agreement_id = agreement.id
                    existing.source_text = point_text[:200] if point_text else None
                else:
                    override = CalendarRouteOverride(
                        date=override_date,
                        contract_id=contract_id,
                        route=route,
                        treat_as=treat_as,
                        source_agreement_id=agreement.id,
                        source_text=point_text[:200] if point_text else None,
                    )
                    db.add(override)
                    added_overrides += 1

    agreement.status = AgreementStatus.APPLIED
    db.commit()

    # Пересчитываем этапы контракта после применения ДС
    stages_updated = recalculate_stages(
        db,
        contract_id,
        source_agreement_id=agreement.id,
    )

    return {
        "success": True,
        "message": f"ДС применён: добавлено {added_params} параметров маршрутов, {added_overrides} переопределений, пересчитано {stages_updated} этапов",
        "seasonal_changes": seasonal_changes,
    }


@router.post("/{agreement_id}/unapply")
async def unapply_agreement(agreement_id: int, db: Session = Depends(get_db)):
    """
    Отменить применение ДС — удалить связанные route_params и calendar_overrides.
    """
    agreement = db.query(Agreement).filter(Agreement.id == agreement_id).first()
    if not agreement:
        raise HTTPException(404, "ДС не найден")

    if agreement.status != AgreementStatus.APPLIED:
        raise HTTPException(400, "ДС не применён")

    # Сначала явно удаляем trips (bulk delete не триггерит ORM cascade)
    param_ids = [p.id for p in db.query(RouteParams.id).filter(
        RouteParams.source_agreement_id == agreement_id
    ).all()]
    if param_ids:
        db.query(RouteTrips).filter(
            RouteTrips.route_params_id.in_(param_ids)
        ).delete(synchronize_session=False)

    # Удаляем route_params
    deleted_params = db.query(RouteParams).filter(
        RouteParams.source_agreement_id == agreement_id
    ).delete(synchronize_session=False)

    # Удаляем calendar_overrides
    deleted_overrides = db.query(CalendarRouteOverride).filter(
        CalendarRouteOverride.source_agreement_id == agreement_id
    ).delete(synchronize_session=False)

    # Удаляем calculated_stages
    deleted_stages = db.query(CalculatedStage).filter(
        CalculatedStage.source_agreement_id == agreement_id
    ).delete(synchronize_session=False)

    agreement.status = AgreementStatus.CHECKED
    db.commit()

    return {
        "success": True,
        "message": f"Применение отменено: удалено {deleted_params} параметров, {deleted_overrides} переопределений, {deleted_stages} этапов",
    }


@router.delete("/{agreement_id}")
async def delete_agreement(agreement_id: int, force: bool = False, db: Session = Depends(get_db)):
    """Удалить ДС. С force=True удаляет и применённые ДС."""
    agreement = db.query(Agreement).filter(Agreement.id == agreement_id).first()
    if not agreement:
        raise HTTPException(404, "ДС не найден")

    contract_id = agreement.contract_id

    if agreement.status == AgreementStatus.APPLIED:
        if not force:
            raise HTTPException(400, "ДС применён. Используйте force=true для удаления")
        # Сначала явно удаляем trips (bulk delete не триггерит ORM cascade)
        param_ids = [p.id for p in db.query(RouteParams.id).filter(
            RouteParams.source_agreement_id == agreement_id
        ).all()]
        if param_ids:
            db.query(RouteTrips).filter(
                RouteTrips.route_params_id.in_(param_ids)
            ).delete(synchronize_session=False)
        # Удаляем связанные данные
        db.query(RouteParams).filter(RouteParams.source_agreement_id == agreement_id).delete(synchronize_session=False)
        db.query(CalendarRouteOverride).filter(CalendarRouteOverride.source_agreement_id == agreement_id).delete(synchronize_session=False)
        db.query(CalculatedStage).filter(CalculatedStage.source_agreement_id == agreement_id).delete(synchronize_session=False)

    db.delete(agreement)
    db.flush()

    # Удаляем осиротевшие overrides и stages (source_agreement_id != NULL, но ДС уже нет)
    # Ручные overrides (source_agreement_id IS NULL) не трогаем
    existing_agreement_ids = [a.id for a in db.query(Agreement.id).filter(
        Agreement.contract_id == contract_id
    ).all()]
    if existing_agreement_ids:
        db.query(CalendarRouteOverride).filter(
            CalendarRouteOverride.contract_id == contract_id,
            CalendarRouteOverride.source_agreement_id.isnot(None),
            ~CalendarRouteOverride.source_agreement_id.in_(existing_agreement_ids),
        ).delete(synchronize_session=False)
        db.query(CalculatedStage).filter(
            CalculatedStage.contract_id == contract_id,
            CalculatedStage.source_agreement_id.isnot(None),
            ~CalculatedStage.source_agreement_id.in_(existing_agreement_ids),
        ).delete(synchronize_session=False)
    else:
        # Нет ДС вообще — удаляем всё кроме ручных записей
        db.query(CalendarRouteOverride).filter(
            CalendarRouteOverride.contract_id == contract_id,
            CalendarRouteOverride.source_agreement_id.isnot(None),
        ).delete(synchronize_session=False)
        db.query(CalculatedStage).filter(
            CalculatedStage.contract_id == contract_id,
        ).delete(synchronize_session=False)

    db.commit()

    return {"success": True}
