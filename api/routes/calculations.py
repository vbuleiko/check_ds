"""API для расчётов."""
from datetime import date
from decimal import Decimal
from typing import Optional
from urllib.parse import quote

from fastapi import APIRouter, Depends, HTTPException, Query
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session

from db.database import get_db
from db.models import Contract, Agreement, RouteParams, CalendarRouteOverride
from core.calculator.kilometers import calculate_route_period, calculate_contract_period, PeriodCalculation, DayCalculation
from core.calculator.compare import compare_agreement
from core.calculator.price import get_coefficients_for_date, get_capacities, preload_price_data
from core.calculator.export_excel import (
    generate_monthly_volumes_excel,
    generate_monthly_volumes_rub_excel,
    generate_monthly_volumes_combined_excel,
    generate_quarterly_volumes_excel,
    get_contract_routes,
    normalize_route,
)

router = APIRouter()

MAX_DATE_RANGE_DAYS = 3660  # ~10 лет


def _parse_and_validate_dates(date_from: str, date_to: str) -> tuple[date, date]:
    """Парсит и валидирует пару дат из строк."""
    try:
        d_from = date.fromisoformat(date_from)
        d_to = date.fromisoformat(date_to)
    except ValueError:
        raise HTTPException(400, "Неверный формат даты (ожидается YYYY-MM-DD)")

    if d_from > d_to:
        raise HTTPException(400, "Дата начала не может быть позже даты окончания")

    if (d_to - d_from).days > MAX_DATE_RANGE_DAYS:
        raise HTTPException(400, f"Слишком большой диапазон дат (максимум {MAX_DATE_RANGE_DAYS} дней)")

    return d_from, d_to


def _merge_route_calculations(route: str, calcs: list[PeriodCalculation]) -> PeriodCalculation:
    """Объединяет расчёты нескольких вариантов маршрута (например, 315 и 315.)."""
    if len(calcs) == 1:
        calcs[0].route = route
        return calcs[0]

    merged = PeriodCalculation(
        route=route,
        date_from=min(c.date_from for c in calcs),
        date_to=max(c.date_to for c in calcs),
    )

    # Суммируем дни по датам
    days_by_date: dict[date, DayCalculation] = {}
    for c in calcs:
        for day in c.days:
            if day.date not in days_by_date:
                days_by_date[day.date] = DayCalculation(
                    date=day.date,
                    day_type=day.day_type,
                    route_params_id=day.route_params_id,
                )
            existing = days_by_date[day.date]
            existing.forward_trips += day.forward_trips
            existing.reverse_trips += day.reverse_trips
            existing.forward_km += day.forward_km
            existing.reverse_km += day.reverse_km
    merged.days = sorted(days_by_date.values(), key=lambda d: d.date)

    # Суммируем по типам дней
    for c in calcs:
        for dt, stats in c.by_day_type.items():
            if dt not in merged.by_day_type:
                merged.by_day_type[dt] = {k: 0 if isinstance(v, (int, float)) else v for k, v in stats.items()}
            for k, v in stats.items():
                if isinstance(v, (int, float)):
                    merged.by_day_type[dt][k] = merged.by_day_type[dt].get(k, 0) + v

    # Объединяем сегменты
    merged.segments = [s for c in calcs for s in c.segments]

    return merged


@router.get("/route/{contract_number}/{route}")
async def calculate_route(
    contract_number: str,
    route: str,
    date_from: str = Query(..., description="Начало периода YYYY-MM-DD"),
    date_to: str = Query(..., description="Конец периода YYYY-MM-DD"),
    db: Session = Depends(get_db)
):
    """
    Рассчитать км по маршруту за период.
    """
    contract = db.query(Contract).filter(Contract.number == contract_number).first()
    if not contract:
        raise HTTPException(404, "Контракт не найден")

    d_from, d_to = _parse_and_validate_dates(date_from, date_to)

    # Ищем все варианты маршрута (например, 315 и 315.)
    norm = normalize_route(route)
    all_routes = get_contract_routes(db, contract.id, normalize=False)
    variants = [r for r in all_routes if normalize_route(r) == norm]
    if not variants:
        variants = [route]

    calcs = [calculate_route_period(db, contract.id, r, d_from, d_to) for r in variants]
    result = _merge_route_calculations(norm, calcs)

    result_dict = result.to_dict()

    # Рассчитываем стоимость в рублях
    try:
        preload_price_data(contract_number)
        capacities = get_capacities(contract_number)
        coefficients = get_coefficients_for_date(d_to, contract_number)

        if capacities and coefficients:
            cap = capacities.get(norm)
            coef = coefficients.get(norm)
            if cap and coef:
                fwd_km = Decimal(str(round(result.total_forward_km, 2)))
                rev_km = Decimal(str(round(result.total_reverse_km, 2)))
                cap_d = Decimal(str(cap))
                fwd_rub = round(float(fwd_km * cap_d * coef), 2)
                rev_rub = round(float(rev_km * cap_d * coef), 2)
                result_dict.update({
                    "price_available": True,
                    "total_rub": round(fwd_rub + rev_rub, 2),
                    "forward_rub": fwd_rub,
                    "reverse_rub": rev_rub,
                })
            else:
                result_dict.update({"price_available": False, "total_rub": None, "forward_rub": None, "reverse_rub": None})
        else:
            result_dict.update({"price_available": False, "total_rub": None, "forward_rub": None, "reverse_rub": None})
    except Exception:
        result_dict.update({"price_available": False, "total_rub": None, "forward_rub": None, "reverse_rub": None})

    return result_dict


@router.get("/route-auto/{route}")
async def calculate_route_auto(
    route: str,
    date_from: str = Query(..., description="Начало периода YYYY-MM-DD"),
    date_to: str = Query(..., description="Конец периода YYYY-MM-DD"),
    db: Session = Depends(get_db)
):
    """
    Рассчитать км по маршруту за период — контракт определяется автоматически.
    """
    norm = normalize_route(route)
    variants = [norm, norm + "."]

    contracts = (
        db.query(Contract)
        .join(RouteParams, RouteParams.contract_id == Contract.id)
        .filter(RouteParams.route.in_(variants))
        .distinct()
        .all()
    )
    if not contracts:
        raise HTTPException(404, f"Маршрут {route} не найден ни в одном контракте")

    contract = contracts[0]
    contract_number = contract.number

    d_from, d_to = _parse_and_validate_dates(date_from, date_to)

    all_routes = get_contract_routes(db, contract.id, normalize=False)
    route_variants = [r for r in all_routes if normalize_route(r) == norm]
    if not route_variants:
        route_variants = [route]

    calcs = [calculate_route_period(db, contract.id, r, d_from, d_to) for r in route_variants]
    result = _merge_route_calculations(norm, calcs)
    result_dict = result.to_dict()
    result_dict["contract_number"] = contract_number

    try:
        preload_price_data(contract_number)
        capacities = get_capacities(contract_number)
        coefficients = get_coefficients_for_date(d_to, contract_number)
        if capacities and coefficients:
            cap = capacities.get(norm)
            coef = coefficients.get(norm)
            if cap and coef:
                fwd_km = Decimal(str(round(result.total_forward_km, 2)))
                rev_km = Decimal(str(round(result.total_reverse_km, 2)))
                cap_d = Decimal(str(cap))
                fwd_rub = round(float(fwd_km * cap_d * coef), 2)
                rev_rub = round(float(rev_km * cap_d * coef), 2)
                result_dict.update({"price_available": True, "total_rub": round(fwd_rub + rev_rub, 2), "forward_rub": fwd_rub, "reverse_rub": rev_rub})
            else:
                result_dict.update({"price_available": False, "total_rub": None, "forward_rub": None, "reverse_rub": None})
        else:
            result_dict.update({"price_available": False, "total_rub": None, "forward_rub": None, "reverse_rub": None})
    except Exception:
        result_dict.update({"price_available": False, "total_rub": None, "forward_rub": None, "reverse_rub": None})

    # Календарные переопределения, влияющие на расчёт
    overrides = (
        db.query(CalendarRouteOverride)
        .filter(
            CalendarRouteOverride.contract_id == contract.id,
            CalendarRouteOverride.route.in_(route_variants),
            CalendarRouteOverride.date >= d_from,
            CalendarRouteOverride.date <= d_to,
        )
        .order_by(CalendarRouteOverride.date)
        .all()
    )
    if overrides:
        # Группируем по source_agreement_id + treat_as для компактного отображения
        groups: dict[tuple, dict] = {}
        for o in overrides:
            key = (o.source_agreement_id, o.treat_as, o.source_text)
            if key not in groups:
                ds_num = None
                if o.source_agreement_id:
                    agr = db.query(Agreement.number).filter(Agreement.id == o.source_agreement_id).first()
                    ds_num = agr[0] if agr else None
                groups[key] = {
                    "ds_number": ds_num,
                    "treat_as": o.treat_as,
                    "source_text": o.source_text,
                    "dates": [],
                }
            groups[key]["dates"].append(str(o.date))

        # Сжимаем даты в диапазоны
        result_dict["calendar_overrides"] = [
            {
                "ds_number": g["ds_number"],
                "treat_as": g["treat_as"],
                "source_text": g["source_text"],
                "dates": _compress_dates(g["dates"]),
                "count": len(g["dates"]),
            }
            for g in groups.values()
        ]

    return result_dict


def _compress_dates(date_strings: list[str]) -> str:
    """Сжимает список дат в читаемые диапазоны: '01.01–11.01.2026'."""
    from datetime import datetime as dt
    dates = sorted(dt.strptime(d, "%Y-%m-%d").date() for d in date_strings)
    if not dates:
        return ""
    ranges = []
    start = dates[0]
    prev = dates[0]
    for d in dates[1:]:
        if (d - prev).days == 1:
            prev = d
        else:
            ranges.append((start, prev))
            start = d
            prev = d
    ranges.append((start, prev))

    parts = []
    for s, e in ranges:
        if s == e:
            parts.append(s.strftime("%d.%m.%Y"))
        else:
            parts.append(f"{s.strftime('%d.%m')}–{e.strftime('%d.%m.%Y')}")
    return ", ".join(parts)


@router.get("/contract/{contract_number}")
async def calculate_contract(
    contract_number: str,
    date_from: str = Query(..., description="Начало периода YYYY-MM-DD"),
    date_to: str = Query(..., description="Конец периода YYYY-MM-DD"),
    routes: str = Query(default=None, description="Маршруты через запятую"),
    db: Session = Depends(get_db)
):
    """
    Рассчитать км по всем маршрутам контракта за период.
    """
    contract = db.query(Contract).filter(Contract.number == contract_number).first()
    if not contract:
        raise HTTPException(404, "Контракт не найден")

    d_from, d_to = _parse_and_validate_dates(date_from, date_to)

    route_list = None
    if routes:
        route_list = [r.strip() for r in routes.split(",")]

    results = calculate_contract_period(db, contract.id, d_from, d_to, route_list)

    # Формируем ответ (округляем каждый маршрут до 2 знаков перед суммированием)
    total_km = sum(round(r.total_km, 2) for r in results.values())
    total_trips = sum(r.total_trips for r in results.values())

    return {
        "contract_number": contract_number,
        "date_from": date_from,
        "date_to": date_to,
        "total_km": round(total_km, 2),
        "total_trips": total_trips,
        "routes_count": len(results),
        "routes": {route: calc.to_dict() for route, calc in results.items()},
    }


@router.get("/agreement/{agreement_id}/compare")
async def compare_agreement_endpoint(
    agreement_id: int,
    calculation_end_date: str = Query(default=None),
    db: Session = Depends(get_db)
):
    """
    Сравнить объёмы до и после применения ДС.
    """
    agreement = db.query(Agreement).filter(Agreement.id == agreement_id).first()
    if not agreement:
        raise HTTPException(404, "ДС не найден")

    end_date = None
    if calculation_end_date:
        try:
            end_date = date.fromisoformat(calculation_end_date)
        except ValueError:
            raise HTTPException(400, "Неверный формат даты")

    result = compare_agreement(db, agreement, end_date)

    # Сохраняем результаты в ДС
    agreement.total_km_before = result.total_km_before
    agreement.total_km_after = result.total_km_after
    agreement.calculation_details = result.to_dict()
    db.commit()

    return result.to_dict()


@router.get("/monthly-plan/{contract_number}")
async def monthly_plan(
    contract_number: str,
    year: int,
    month: int,
    db: Session = Depends(get_db)
):
    """
    План на месяц по контракту.
    """
    contract = db.query(Contract).filter(Contract.number == contract_number).first()
    if not contract:
        raise HTTPException(404, "Контракт не найден")

    # Определяем период
    d_from = date(year, month, 1)
    if month == 12:
        d_to = date(year, 12, 31)
    else:
        d_to = date(year, month + 1, 1)
        from datetime import timedelta
        d_to = d_to - timedelta(days=1)

    results = calculate_contract_period(db, contract.id, d_from, d_to)

    # Формируем план
    routes_plan = []
    for route, calc in sorted(results.items()):
        routes_plan.append({
            "route": route,
            "total_km": round(calc.total_km, 2),
            "forward_km": round(calc.total_forward_km, 2),
            "reverse_km": round(calc.total_reverse_km, 2),
            "total_trips": calc.total_trips,
            "forward_trips": calc.total_forward_trips,
            "reverse_trips": calc.total_reverse_trips,
        })

    total_km = sum(r["total_km"] for r in routes_plan)
    total_trips = sum(r["total_trips"] for r in routes_plan)

    return {
        "contract_number": contract_number,
        "year": year,
        "month": month,
        "period": f"{d_from} - {d_to}",
        "total_km": round(total_km, 2),
        "total_trips": total_trips,
        "routes": routes_plan,
    }


@router.get("/export/{contract_number}")
async def export_monthly_volumes(
    contract_number: str,
    start_year: int = Query(default=2026, description="Начальный год"),
    start_month: int = Query(default=2, description="Начальный месяц"),
    end_year: int = Query(default=2028, description="Конечный год"),
    end_month: int = Query(default=7, description="Конечный месяц"),
    ds_number: str = Query(default=None, description="Номер ДС для заголовка"),
    db: Session = Depends(get_db)
):
    """
    Экспорт помесячных объёмов работ в Excel.

    Генерирует файл по образцу Приложения №12.
    По умолчанию: февраль 2026 - июль 2028.
    """
    contract = db.query(Contract).filter(Contract.number == contract_number).first()
    if not contract:
        raise HTTPException(404, "Контракт не найден")

    output = generate_monthly_volumes_excel(
        session=db,
        contract_id=contract.id,
        start_year=start_year,
        start_month=start_month,
        end_year=end_year,
        end_month=end_month,
        ds_number=ds_number,
    )

    filename = f"Объёмы_ГК{contract_number}"
    if ds_number:
        filename += f"_ДС{ds_number}"
    filename += f"_{start_month:02d}.{start_year}-{end_month:02d}.{end_year}.xlsx"

    # URL-encode filename for Content-Disposition header
    filename_encoded = quote(filename)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{filename_encoded}"}
    )


@router.get("/export-quarterly/{contract_number}")
async def export_quarterly_volumes(
    contract_number: str,
    start_year: int = Query(default=2022, description="Начальный год"),
    start_quarter: int = Query(default=2, description="Начальный квартал"),
    end_year: int = Query(default=2028, description="Конечный год"),
    end_quarter: int = Query(default=3, description="Конечный квартал"),
    ds_number: str = Query(default=None, description="Номер ДС для заголовка"),
    db: Session = Depends(get_db)
):
    """
    Экспорт поквартальных объёмов работ в Excel.

    Формат как в исходном Приложении №12.
    """
    contract = db.query(Contract).filter(Contract.number == contract_number).first()
    if not contract:
        raise HTTPException(404, "Контракт не найден")

    output = generate_quarterly_volumes_excel(
        session=db,
        contract_id=contract.id,
        start_year=start_year,
        start_quarter=start_quarter,
        end_year=end_year,
        end_quarter=end_quarter,
        ds_number=ds_number,
    )

    filename = f"Объёмы_квартальные_ГК{contract_number}"
    if ds_number:
        filename += f"_ДС{ds_number}"
    filename += f".xlsx"

    filename_encoded = quote(filename)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{filename_encoded}"}
    )


@router.get("/export-rub/{contract_number}")
async def export_monthly_volumes_rub(
    contract_number: str,
    start_year: int = Query(default=2026, description="Начальный год"),
    start_month: int = Query(default=2, description="Начальный месяц"),
    end_year: int = Query(default=2028, description="Конечный год"),
    end_month: int = Query(default=7, description="Конечный месяц"),
    ds_number: str = Query(default=None, description="Номер ДС для заголовка"),
    db: Session = Depends(get_db)
):
    """Экспорт помесячных объёмов работ в рублях в Excel."""
    contract = db.query(Contract).filter(Contract.number == contract_number).first()
    if not contract:
        raise HTTPException(404, "Контракт не найден")

    output = generate_monthly_volumes_rub_excel(
        session=db,
        contract_id=contract.id,
        contract_number=contract_number,
        start_year=start_year,
        start_month=start_month,
        end_year=end_year,
        end_month=end_month,
        ds_number=ds_number,
    )

    filename = f"Объёмы_руб_ГК{contract_number}"
    if ds_number:
        filename += f"_ДС{ds_number}"
    filename += f"_{start_month:02d}.{start_year}-{end_month:02d}.{end_year}.xlsx"
    filename_encoded = quote(filename)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{filename_encoded}"}
    )


@router.get("/export-combined/{contract_number}")
async def export_monthly_volumes_combined(
    contract_number: str,
    start_year: int = Query(default=2026, description="Начальный год"),
    start_month: int = Query(default=2, description="Начальный месяц"),
    end_year: int = Query(default=2028, description="Конечный год"),
    end_month: int = Query(default=7, description="Конечный месяц"),
    ds_number: str = Query(default=None, description="Номер ДС для заголовка"),
    db: Session = Depends(get_db)
):
    """Экспорт помесячных объёмов работ (км и руб.) в один Excel файл с двумя листами."""
    contract = db.query(Contract).filter(Contract.number == contract_number).first()
    if not contract:
        raise HTTPException(404, "Контракт не найден")

    output = generate_monthly_volumes_combined_excel(
        session=db,
        contract_id=contract.id,
        contract_number=contract_number,
        start_year=start_year,
        start_month=start_month,
        end_year=end_year,
        end_month=end_month,
        ds_number=ds_number,
    )

    filename = f"Объёмы_км_и_руб_ГК{contract_number}"
    if ds_number:
        filename += f"_ДС{ds_number}"
    filename += f"_{start_month:02d}.{start_year}-{end_month:02d}.{end_year}.xlsx"
    filename_encoded = quote(filename)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{filename_encoded}"}
    )


@router.get("/routes/{contract_number}")
async def get_routes(
    contract_number: str,
    db: Session = Depends(get_db)
):
    """
    Получить список маршрутов контракта.
    """
    contract = db.query(Contract).filter(Contract.number == contract_number).first()
    if not contract:
        raise HTTPException(404, "Контракт не найден")

    routes = get_contract_routes(db, contract.id)

    return {
        "contract_number": contract_number,
        "routes_count": len(routes),
        "routes": routes,
    }
