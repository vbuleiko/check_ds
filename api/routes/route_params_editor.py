"""API для редактирования параметров маршрутов."""
from datetime import date
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from sqlalchemy.orm import Session

from sqlalchemy import func

from db.database import get_db
from db.models import Contract, RouteParams, RouteTrips, SeasonType, CalculatedStage, StageStatus
from core.constants import get_weekdays_for_type_extended

router = APIRouter()


@router.get("/routes")
async def list_routes(db: Session = Depends(get_db)):
    """Список уникальных маршрутов по всем контрактам."""
    rows = (
        db.query(RouteParams.route)
        .distinct()
        .order_by(RouteParams.route)
        .all()
    )
    return {"routes": [r[0] for r in rows]}


@router.get("/params")
async def get_params(route: str, db: Session = Depends(get_db)):
    """Все записи RouteParams для маршрута по всем контрактам (с рейсами).
    Ищет как точное совпадение, так и вариант с точкой на конце (315 и 315.).
    """
    base = route.rstrip(".")
    variants = {base, base + "."}
    params_list = (
        db.query(RouteParams)
        .filter(RouteParams.route.in_(variants))
        .order_by(RouteParams.date_from)
        .all()
    )

    # Загружаем номера контрактов одним запросом
    contract_ids = {p.contract_id for p in params_list}
    contracts = {c.id: c.number for c in db.query(Contract).filter(Contract.id.in_(contract_ids)).all()}

    # Максимальная дата_to закрытых этапов по каждому контракту
    closed_rows = (
        db.query(CalculatedStage.contract_id, func.max(CalculatedStage.date_to))
        .filter(
            CalculatedStage.contract_id.in_(contract_ids),
            CalculatedStage.status == StageStatus.CLOSED,
        )
        .group_by(CalculatedStage.contract_id)
        .all()
    )
    max_closed: dict[int, date] = {cid: dt for cid, dt in closed_rows}

    result = []
    for p in params_list:
        closed_until = max_closed.get(p.contract_id)
        if p.date_to is not None and closed_until is not None and p.date_to <= closed_until:
            is_active = False
        else:
            is_active = True

        result.append({
            "id": p.id,
            "contract_number": contracts.get(p.contract_id, "?"),
            "route": p.route,
            "date_from": str(p.date_from),
            "date_to": str(p.date_to) if p.date_to else None,
            "season": p.season.value,
            "length_forward": p.length_forward,
            "length_reverse": p.length_reverse,
            "source_appendix": p.source_appendix,
            "is_active": is_active,
            "trips": [
                {
                    "id": t.id,
                    "day_type_name": t.day_type_name,
                    "weekdays": t.weekdays,
                    "forward_number": t.forward_number,
                    "reverse_number": t.reverse_number,
                }
                for t in p.trips
            ],
        })
    return {"params": result}


class UpdateParamsRequest(BaseModel):
    route: Optional[str] = None
    date_from: Optional[str] = None
    date_to: Optional[str] = None
    season: Optional[str] = None
    length_forward: Optional[float] = None
    length_reverse: Optional[float] = None
    source_appendix: Optional[str] = None


@router.put("/params/{param_id}")
async def update_params(
    param_id: int,
    data: UpdateParamsRequest,
    db: Session = Depends(get_db),
):
    """Обновить поля RouteParams."""
    p = db.query(RouteParams).filter(RouteParams.id == param_id).first()
    if not p:
        raise HTTPException(404, "Запись не найдена")

    if data.route is not None:
        p.route = data.route
    if data.date_from is not None:
        p.date_from = date.fromisoformat(data.date_from)
    if data.date_to is not None:
        p.date_to = date.fromisoformat(data.date_to) if data.date_to else None
    if data.season is not None:
        p.season = SeasonType(data.season)
    if data.length_forward is not None:
        p.length_forward = data.length_forward
    if data.length_reverse is not None:
        p.length_reverse = data.length_reverse
    if data.source_appendix is not None:
        p.source_appendix = data.source_appendix

    db.commit()
    return {"success": True}


class UpdateTripRequest(BaseModel):
    day_type_name: Optional[str] = None
    forward_number: Optional[int] = None
    reverse_number: Optional[int] = None


@router.put("/trips/{trip_id}")
async def update_trip(
    trip_id: int,
    data: UpdateTripRequest,
    db: Session = Depends(get_db),
):
    """Обновить запись RouteTrips."""
    t = db.query(RouteTrips).filter(RouteTrips.id == trip_id).first()
    if not t:
        raise HTTPException(404, "Запись не найдена")

    if data.day_type_name is not None:
        t.day_type_name = data.day_type_name
        weekdays = get_weekdays_for_type_extended(data.day_type_name)
        if weekdays:
            t.weekdays = weekdays
    if data.forward_number is not None:
        t.forward_number = data.forward_number
    if data.reverse_number is not None:
        t.reverse_number = data.reverse_number

    db.commit()
    return {"success": True}


class CreateTripRequest(BaseModel):
    day_type_name: str
    forward_number: int = 0
    reverse_number: int = 0


@router.post("/params/{param_id}/trips")
async def create_trip(
    param_id: int,
    data: CreateTripRequest,
    db: Session = Depends(get_db),
):
    """Добавить тип дня (рейс) к записи RouteParams."""
    p = db.query(RouteParams).filter(RouteParams.id == param_id).first()
    if not p:
        raise HTTPException(404, "Запись не найдена")

    weekdays = get_weekdays_for_type_extended(data.day_type_name) or []
    t = RouteTrips(
        route_params_id=param_id,
        day_type_name=data.day_type_name,
        weekdays=weekdays,
        forward_number=data.forward_number,
        reverse_number=data.reverse_number,
    )
    db.add(t)
    db.commit()
    db.refresh(t)
    return {
        "id": t.id,
        "day_type_name": t.day_type_name,
        "weekdays": t.weekdays,
        "forward_number": t.forward_number,
        "reverse_number": t.reverse_number,
    }


@router.delete("/trips/{trip_id}")
async def delete_trip(trip_id: int, db: Session = Depends(get_db)):
    """Удалить тип дня (рейс)."""
    t = db.query(RouteTrips).filter(RouteTrips.id == trip_id).first()
    if not t:
        raise HTTPException(404, "Запись не найдена")
    db.delete(t)
    db.commit()
    return {"success": True}


@router.delete("/params/{param_id}")
async def delete_params(param_id: int, db: Session = Depends(get_db)):
    """Удалить запись RouteParams вместе со всеми рейсами."""
    p = db.query(RouteParams).filter(RouteParams.id == param_id).first()
    if not p:
        raise HTTPException(404, "Запись не найдена")
    db.delete(p)
    db.commit()
    return {"success": True}
