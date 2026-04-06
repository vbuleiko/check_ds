"""API для вкладки «Проверка таблиц»."""
from typing import Optional
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from sqlalchemy.orm import Session

from sqlalchemy.orm import joinedload

from db.database import get_db
from db.models import Agreement, AgreementReference, AgreementStatus, Contract
from core.calculator.price import invalidate_cache
from core.constants import ROUTE_SEASON_PERIODS
from core.checker.internal import _parse_season_date, _fmt_season_date

from api.routes.table_checks_logic import (
    ETAPY_SROKI_HEADER_MAP,
    check_etapy_avans_table,
    check_etapy_sroki_table,
    check_finansirovanie_table,
    check_km_by_routes,
    check_km_total_vs_probeg,
    check_price_change,
    check_raschet_table,
    check_vysv_closed_stage,
    find_all_previous_agreements,
    find_previous_agreement,
)

router = APIRouter()


# =============================================================================
# Эндпоинты
# =============================================================================



@router.post("/cache/clear")
async def clear_price_cache():
    """Очищает кэш коэффициентов и вместимостей (перезагрузится при следующем запросе)."""
    invalidate_cache()
    return {"ok": True, "message": "Кэш очищен"}


@router.get("/agreements")
async def list_agreements_for_checks(
    contract_number: Optional[str] = None,
    db: Session = Depends(get_db),
):
    """Список ДС для выбора на вкладке «Проверка таблиц» (без черновиков, отсортированы по убыванию номера)."""
    query = (
        db.query(Agreement)
        .options(joinedload(Agreement.contract))
        .filter(Agreement.status != AgreementStatus.DRAFT)
    )
    if contract_number:
        contract = db.query(Contract).filter(Contract.number == contract_number).first()
        if contract:
            query = query.filter(Agreement.contract_id == contract.id)

    agreements = query.all()

    result = [
        {
            "id": a.id,
            "number": a.number,
            "number_int": int(a.number) if a.number.isdigit() else -1,
            "contract_number": a.contract.number if a.contract else None,
            "contract_id": a.contract_id,
            "status": a.status,
        }
        for a in agreements
    ]

    result.sort(key=lambda x: (x["contract_id"], x["number_int"]), reverse=True)
    return result


@router.get("/references")
async def list_references(
    contract_number: Optional[str] = None,
    db: Session = Depends(get_db),
):
    """Список эталонных записей."""
    query = db.query(AgreementReference)
    if contract_number:
        contract = db.query(Contract).filter(Contract.number == contract_number).first()
        if contract:
            query = query.filter(AgreementReference.contract_id == contract.id)

    refs = query.all()
    result = []
    for r in refs:
        contract = db.query(Contract).filter(Contract.id == r.contract_id).first()
        result.append({
            "id": r.id,
            "contract_id": r.contract_id,
            "contract_number": contract.number if contract else None,
            "reference_ds_number": r.reference_ds_number,
            "initial_km": r.initial_km,
            "probeg_etapy": r.probeg_etapy,
            "sum_price": r.sum_price,
            "note": r.note,
            "created_at": str(r.created_at),
        })
    return result


class ReferenceCreateRequest(BaseModel):
    contract_number: str
    reference_ds_number: str
    initial_km: Optional[float] = None
    probeg_etapy: Optional[float] = None
    sum_price: Optional[float] = None
    note: Optional[str] = None


@router.post("/references")
async def create_or_update_reference(
    body: ReferenceCreateRequest,
    db: Session = Depends(get_db),
):
    """Создать или обновить эталонную запись."""
    contract = db.query(Contract).filter(Contract.number == body.contract_number).first()
    if not contract:
        raise HTTPException(status_code=404, detail=f"Контракт {body.contract_number} не найден")

    ref = db.query(AgreementReference).filter(
        AgreementReference.contract_id == contract.id,
        AgreementReference.reference_ds_number == body.reference_ds_number,
    ).first()

    if ref:
        ref.initial_km = body.initial_km
        ref.probeg_etapy = body.probeg_etapy
        ref.sum_price = body.sum_price
        ref.note = body.note
    else:
        ref = AgreementReference(
            contract_id=contract.id,
            reference_ds_number=body.reference_ds_number,
            initial_km=body.initial_km,
            probeg_etapy=body.probeg_etapy,
            sum_price=body.sum_price,
            note=body.note,
        )
        db.add(ref)

    db.commit()
    db.refresh(ref)
    return {"ok": True, "id": ref.id}


@router.delete("/references/{ref_id}")
async def delete_reference(ref_id: int, db: Session = Depends(get_db)):
    """Удалить эталонную запись."""
    ref = db.query(AgreementReference).filter(AgreementReference.id == ref_id).first()
    if not ref:
        raise HTTPException(status_code=404, detail="Запись не найдена")
    db.delete(ref)
    db.commit()
    return {"ok": True}


def _get_table_checks_vysv(agreement, json_data: dict, contract, db: Session) -> dict:
    """Возвращает данные проверок таблиц для ДС на высвобождение."""
    general = json_data.get("general", {})
    contract_number = contract.number if contract else None
    contract_id = agreement.contract_id

    # Сырые таблицы из парсера
    table_etapy_sroki = json_data.get("stages_table1_raw", [])
    table_etapy_avans = json_data.get("stages_table2_raw", [])
    table_finansirovanie = json_data.get("stages_finansirovanie_raw", [])

    # Ищем предыдущие ДС (нужны для проверки цены)
    all_prev_agreements, all_prev_references = [], []
    prev_agreement, prev_reference = None, None
    if contract_id and agreement.number:
        prev_agreement, prev_reference = find_previous_agreement(db, contract_id, agreement.number)
        all_prev_agreements, all_prev_references = find_all_previous_agreements(db, contract_id, agreement.number)

    etapy_headers = [
        ETAPY_SROKI_HEADER_MAP.get(h, h)
        for h in (table_etapy_sroki[0] if table_etapy_sroki else [])
    ]
    etapy_avans_headers = [
        ETAPY_SROKI_HEADER_MAP.get(h, h)
        for h in (table_etapy_avans[0] if table_etapy_avans else [])
    ]

    # Проверки таблицы этапов (сроки)
    if table_etapy_sroki and contract_number and contract_id:
        checks_etapy = check_etapy_sroki_table(
            table_etapy_sroki,
            db,
            contract_id,
            contract_number,
            prev_agreement,
        )
    else:
        checks_etapy = {"row_checks": [], "itogo_check": None, "price_available": False}

    # Проверки таблицы этапов (с учётом авансов)
    if table_etapy_avans and table_etapy_sroki:
        checks_etapy_avans = check_etapy_avans_table(table_etapy_avans, table_etapy_sroki)
    else:
        checks_etapy_avans = []

    # Проверки таблицы финансирования
    if table_finansirovanie and table_etapy_sroki:
        checks_finansirovanie = check_finansirovanie_table(table_finansirovanie, table_etapy_sroki)
    else:
        checks_finansirovanie = []

    # Проверка закрытого этапа
    if contract_id:
        checks_closed_stage = check_vysv_closed_stage(general, db, contract_id)
    else:
        checks_closed_stage = []

    # Проверка изменения цены (высвобождение тоже может менять цену)
    checks_price_change = check_price_change(json_data, all_prev_agreements, all_prev_references)

    # Информация о предыдущем ДС
    if prev_agreement:
        prev_source = "agreement"
        prev_info = {
            "number": prev_agreement.number,
            "id": prev_agreement.id,
            "status": prev_agreement.status,
        }
    elif prev_reference:
        prev_source = "reference"
        prev_info = {
            "number": prev_reference.reference_ds_number,
            "id": prev_reference.id,
            "initial_km": prev_reference.initial_km,
            "probeg_etapy": prev_reference.probeg_etapy,
            "sum_price": prev_reference.sum_price,
            "note": prev_reference.note,
        }
    else:
        prev_source = None
        prev_info = None

    return {
        "agreement": {
            "id": agreement.id,
            "number": agreement.number,
            "contract_number": contract_number,
            "status": agreement.status,
            "type": "vysvobozhdenie",
        },
        "prev_source": prev_source,
        "prev_info": prev_info,
        "km_routes_not_applicable": True,
        "tables": {
            "table_raschet_izm_objema": {
                "headers": [],
                "rows": [],
                "checks": checks_closed_stage,
                "price_change_checks": checks_price_change,
                "appendix_number": None,
                "ds_number": None,
            },
            "table_etapy_sroki": {
                "headers": etapy_headers,
                "rows": table_etapy_sroki[1:] if len(table_etapy_sroki) > 1 else [],
                "row_checks": checks_etapy["row_checks"],
                "itogo_check": checks_etapy["itogo_check"],
                "price_available": checks_etapy["price_available"],
                "appendix_number": None,
                "ds_number": None,
            },
            "table_finansirovanie_po_godam": {
                "headers": table_finansirovanie[0] if table_finansirovanie else [],
                "rows": table_finansirovanie[1:] if len(table_finansirovanie) > 1 else [],
                "checks": checks_finansirovanie,
                "appendix_number": None,
                "ds_number": None,
            },
            "table_etapy_avans": {
                "headers": etapy_avans_headers,
                "rows": table_etapy_avans[1:] if len(table_etapy_avans) > 1 else [],
                "checks": checks_etapy_avans,
                "appendix_number": None,
                "ds_number": None,
            },
        },
        "km_by_routes": [],
        "km_total_vs_probeg": None,
        "km_appendix_number": None,
    }


@router.get("/{agreement_id}")
async def get_table_checks(agreement_id: int, db: Session = Depends(get_db)):
    """Возвращает данные таблиц и результаты проверок для указанного ДС."""
    agreement = db.query(Agreement).filter(Agreement.id == agreement_id).first()
    if not agreement:
        raise HTTPException(status_code=404, detail="ДС не найден")

    if not agreement.json_data:
        raise HTTPException(status_code=400, detail="У этого ДС нет данных JSON")

    contract = db.query(Contract).filter(Contract.id == agreement.contract_id).first()

    json_data = agreement.json_data

    # ДС на высвобождение — отдельная логика
    if json_data.get("type") == "vysvobozhdenie":
        return _get_table_checks_vysv(agreement, json_data, contract, db)

    prev_agreement, prev_reference = find_previous_agreement(
        db, agreement.contract_id, agreement.number
    )
    all_prev_agreements, all_prev_references = find_all_previous_agreements(
        db, agreement.contract_id, agreement.number
    )

    # Определяем источник предыдущих данных (ближайший ДС для отображения)
    if prev_agreement:
        prev_source = "agreement"
        prev_info = {
            "number": prev_agreement.number,
            "id": prev_agreement.id,
            "status": prev_agreement.status,
        }
    elif prev_reference:
        prev_source = "reference"
        prev_info = {
            "number": prev_reference.reference_ds_number,
            "id": prev_reference.id,
            "initial_km": prev_reference.initial_km,
            "probeg_etapy": prev_reference.probeg_etapy,
            "sum_price": prev_reference.sum_price,
            "note": prev_reference.note,
        }
    else:
        prev_source = None
        prev_info = None

    # Таблицы из JSON
    table_raschet = json_data.get("table_raschet_izm_objema", [])
    table_etapy_sroki = json_data.get("table_etapy_sroki", [])
    table_finansirovanie = json_data.get("table_finansirovanie_po_godam", [])
    table_etapy_avans = json_data.get("table_etapy_avans", [])

    # Переименовываем заголовки таблицы этапов
    etapy_headers = [
        ETAPY_SROKI_HEADER_MAP.get(h, h)
        for h in (table_etapy_sroki[0] if table_etapy_sroki else [])
    ]
    etapy_avans_headers = [
        ETAPY_SROKI_HEADER_MAP.get(h, h)
        for h in (table_etapy_avans[0] if table_etapy_avans else [])
    ]

    contract_number = contract.number if contract else "222"

    # Проверки (передаём все предыдущие ДС для поиска по цепочке)
    checks_raschet = check_raschet_table(
        table_raschet, json_data, all_prev_agreements, all_prev_references
    )

    checks_price_change = check_price_change(
        json_data, all_prev_agreements, all_prev_references
    )

    checks_etapy = check_etapy_sroki_table(
        table_etapy_sroki,
        db,
        agreement.contract_id,
        contract_number,
        prev_agreement,
    )

    checks_finansirovanie = check_finansirovanie_table(
        table_finansirovanie,
        table_etapy_sroki,
    )

    checks_etapy_avans = check_etapy_avans_table(
        table_etapy_avans,
        table_etapy_sroki,
    )
    km_data = json_data.get("km_data")

    km_routes_not_applicable = contract_number == "252"
    if km_routes_not_applicable:
        checks_km_routes = []
        km_total_vs_probeg = None
    else:
        checks_km_routes = check_km_by_routes(
            km_data,
            db,
            agreement.contract_id,
            contract_number,
            prev_agreement,
        )
        km_total_vs_probeg = check_km_total_vs_probeg(km_data, json_data.get("general", {}))

    return {
        "agreement": {
            "id": agreement.id,
            "number": agreement.number,
            "contract_number": contract_number,
            "status": agreement.status,
        },
        "prev_source": prev_source,
        "prev_info": prev_info,
        "km_routes_not_applicable": km_routes_not_applicable,
        "tables": {
            "table_raschet_izm_objema": {
                "headers": table_raschet[0] if table_raschet else [],
                "rows": table_raschet[1:] if len(table_raschet) > 1 else [],
                "checks": checks_raschet,
                "price_change_checks": checks_price_change,
                "appendix_number": json_data.get("table_raschet_izm_objema_appendix_number"),
                "ds_number": json_data.get("table_raschet_izm_objema_ds_number"),
            },
            "table_etapy_sroki": {
                "headers": etapy_headers,
                "rows": table_etapy_sroki[1:] if len(table_etapy_sroki) > 1 else [],
                "row_checks": checks_etapy["row_checks"],
                "itogo_check": checks_etapy["itogo_check"],
                "price_available": checks_etapy["price_available"],
                "appendix_number": json_data.get("table_etapy_sroki_appendix_number"),
                "ds_number": json_data.get("table_etapy_sroki_ds_number"),
            },
            "table_finansirovanie_po_godam": {
                "headers": table_finansirovanie[0] if table_finansirovanie else [],
                "rows": table_finansirovanie[1:] if len(table_finansirovanie) > 1 else [],
                "checks": checks_finansirovanie,
                "appendix_number": json_data.get("table_finansirovanie_po_godam_appendix_number"),
                "ds_number": json_data.get("table_finansirovanie_po_godam_ds_number"),
            },
            "table_etapy_avans": {
                "headers": etapy_avans_headers,
                "rows": table_etapy_avans[1:] if len(table_etapy_avans) > 1 else [],
                "checks": checks_etapy_avans,
                "appendix_number": json_data.get("table_etapy_avans_appendix_number"),
                "ds_number": json_data.get("table_etapy_avans_ds_number"),
            },
        },
        "km_by_routes": checks_km_routes,
        "km_total_vs_probeg": km_total_vs_probeg,
        "km_appendix_number": (km_data or {}).get("appendix_number"),
    }


@router.get("/agreement/{agreement_id}/seasonal")
async def get_seasonal_check(agreement_id: int, db: Session = Depends(get_db)):
    """
    Возвращает все сезонные маршруты контракта с проверкой графиков:
    - маршруты с изменениями в этом ДС: сравнение периодов с эталоном
    - маршруты без изменений в этом ДС: "нет изменений в этом ДС"
    """
    from db.models import RouteParams, RouteSeasonConfig, SeasonType

    agreement = db.query(Agreement).filter(Agreement.id == agreement_id).first()
    if not agreement:
        raise HTTPException(404, "ДС не найден")

    json_data = agreement.json_data or {}
    appendices = json_data.get("appendices", {})
    contract_id = agreement.contract_id

    # Загружаем сезонные конфигурации из БД (с fallback на константы)
    db_season_configs = db.query(RouteSeasonConfig).all()
    if db_season_configs:
        season_periods = {
            c.route: {
                "winter": (c.winter_start_month, c.winter_start_day, c.winter_end_month, c.winter_end_day),
                "summer": (c.summer_start_month, c.summer_start_day, c.summer_end_month, c.summer_end_day),
            }
            for c in db_season_configs
        }
    else:
        season_periods = ROUTE_SEASON_PERIODS

    # Собираем маршруты из приложений этого ДС (route_upper -> данные приложения)
    ds_routes: dict[str, dict] = {}
    for appendix_id, app in appendices.items():
        route = (app.get("route") or "").strip().upper()
        if route:
            ds_routes[route] = app

    # Собираем все сезонные маршруты контракта из БД
    db_seasonal_routes: set[str] = set()
    if contract_id:
        db_rows = (
            db.query(RouteParams.route)
            .filter(
                RouteParams.contract_id == contract_id,
                RouteParams.season.in_([SeasonType.WINTER, SeasonType.SUMMER]),
            )
            .distinct()
            .all()
        )
        for (r,) in db_rows:
            normalized = r.strip().upper()
            if season_periods.get(normalized):
                db_seasonal_routes.add(normalized)

    # Итоговый набор: сезонные маршруты из БД + сезонные маршруты из ДС
    all_seasonal = (
        db_seasonal_routes |
        {r for r in ds_routes if season_periods.get(r)}
    )

    # Вспомогательная функция для разбора информации о периоде
    def _period_info(period_data, season_key, ref):
        if not period_data:
            return None
        ds_from = _parse_season_date(period_data.get("date_from"))
        ds_to = _parse_season_date(period_data.get("date_to"))
        ds_range = f"{_fmt_season_date(ds_from)} \u2013 {_fmt_season_date(ds_to)}"
        ref_range = None
        mismatch = False
        if ref and season_key in ref:
            sm, sd, em, ed = ref[season_key]
            ref_from = (sm, sd)
            ref_to = (em, ed)
            ref_range = f"{_fmt_season_date(ref_from)} \u2013 {_fmt_season_date(ref_to)}"
            mismatch = (
                (ds_from is not None and ds_from != ref_from) or
                (ds_to is not None and ds_to != ref_to)
            )
        return {"ds_range": ds_range, "ref_range": ref_range, "mismatch": mismatch}

    # Вспомогательная функция для определения середино-сезонного изменения
    def _seasonal_change(app, ref):
        date_from_str = app.get("date_from") or app.get("date_on")
        if not date_from_str or not ref:
            return None
        parts = date_from_str.split("-")
        if len(parts) != 3:
            return None
        try:
            year, month, day = int(parts[0]), int(parts[1]), int(parts[2])
        except ValueError:
            return None
        if year == 0:
            return None
        is_season_start = any(
            (month, day) == (sm, sd)
            for sm, sd, _, _ in ref.values()
        )
        if is_season_start:
            return None
        for sm, sd, em, ed in ref.values():
            in_season = (
                ((sm, sd) <= (month, day) <= (em, ed)) if sm <= em
                else ((month, day) >= (sm, sd) or (month, day) <= (em, ed))
            )
            if in_season:
                end_year = year + 1 if em < month else year
                return {
                    "year": year,
                    "date_from": f"{day:02d}.{month:02d}.{year}",
                    "date_to": f"{ed:02d}.{em:02d}.{end_year}",
                }
        return None

    rows = []
    for route in sorted(all_seasonal):
        ref = season_periods.get(route)
        app = ds_routes.get(route)

        if app is None:
            # Маршрут есть в контракте, но без изменений в этом ДС
            ref_winter = f"{_fmt_season_date((ref['winter'][0], ref['winter'][1]))} \u2013 {_fmt_season_date((ref['winter'][2], ref['winter'][3]))}" if ref and 'winter' in ref else None
            ref_summer = f"{_fmt_season_date((ref['summer'][0], ref['summer'][1]))} \u2013 {_fmt_season_date((ref['summer'][2], ref['summer'][3]))}" if ref and 'summer' in ref else None
            rows.append({
                "appendix": "\u2014",
                "route": route,
                "winter": {"ds_range": None, "ref_range": ref_winter, "mismatch": False} if ref_winter else None,
                "summer": {"ds_range": None, "ref_range": ref_summer, "mismatch": False} if ref_summer else None,
                "seasonal_change": None,
                "no_period_data": False,
                "no_ds_changes": True,
            })
        else:
            appendix_num = app.get("appendix_num", "?")
            period_winter = app.get("period_winter")
            period_summer = app.get("period_summer")
            rows.append({
                "appendix": f"\u2116{appendix_num}",
                "route": route,
                "winter": _period_info(period_winter, "winter", ref),
                "summer": _period_info(period_summer, "summer", ref),
                "seasonal_change": _seasonal_change(app, ref),
                "no_period_data": period_winter is None and period_summer is None,
                "no_ds_changes": False,
            })

    return {"rows": rows}
