"""API для суммарных данных по контракту."""
from datetime import date
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException, Query
from pydantic import BaseModel
from sqlalchemy.orm import Session

from db.database import get_db
from db.models import Contract, CalculatedStage, StageStatus
from services.contract_summary import get_contract_summary, get_stage_details
from services.stages_calculator import recalculate_contract_stages, has_calculated_stages, recalculate_single_stage
from core.historical_stages import get_contract_config


class StageCloseRequest(BaseModel):
    """Запрос на закрытие этапа."""
    total_km: Optional[float] = None
    total_price: Optional[float] = None

router = APIRouter()


@router.get("/{contract_number}/stages")
async def get_all_stages(
    contract_number: str,
    end_date: str = Query(default=None, description="Конечная дата (YYYY-MM-DD)"),
    db: Session = Depends(get_db)
):
    """
    Получить все этапы контракта.

    Возвращает:
    - Исторические этапы (is_locked=True) — до января 2026 включительно
    - Расчётные этапы (is_locked=False) — с февраля 2026

    Параметры:
    - contract_number: номер контракта (например, "222")
    - end_date: конечная дата расчёта (по умолчанию 2028-07-31)
    """
    end_dt = None
    if end_date:
        try:
            end_dt = date.fromisoformat(end_date)
        except ValueError:
            raise HTTPException(400, "Неверный формат даты (используйте YYYY-MM-DD)")

    summary = get_contract_summary(db, contract_number, end_dt)

    return summary.to_dict()


@router.get("/{contract_number}/totals")
async def get_totals(
    contract_number: str,
    end_date: str = Query(default=None, description="Конечная дата (YYYY-MM-DD)"),
    db: Session = Depends(get_db)
):
    """
    Получить итоговые суммы по контракту.

    Возвращает:
    - historical_km: сумма км по историческим этапам (заблокированы)
    - calculated_km: сумма км по расчётным этапам
    - total_km: общая сумма
    """
    end_dt = None
    if end_date:
        try:
            end_dt = date.fromisoformat(end_date)
        except ValueError:
            raise HTTPException(400, "Неверный формат даты")

    summary = get_contract_summary(db, contract_number, end_dt)

    return {
        "contract_number": contract_number,
        "stages_count": summary.stages_count,
        "historical_stages": get_contract_config(contract_number)["first_stage"] - 1,
        "calculated_stages": summary.stages_count - (get_contract_config(contract_number)["first_stage"] - 1),
        "historical_km": round(summary.historical_km, 2),
        "calculated_km": round(summary.calculated_km, 2),
        "total_km": round(summary.total_km, 2),
        "historical_price": round(summary.historical_price, 2),
        "calculated_price": round(summary.calculated_price, 2),
        "total_price": round(summary.total_price, 2),
    }


@router.get("/{contract_number}/stage/{stage_number}")
async def get_stage(
    contract_number: str,
    stage_number: int,
    db: Session = Depends(get_db)
):
    """
    Получить детали конкретного этапа.

    Для исторических этапов (1-18):
    - Возвращает константные данные
    - routes = null (нет детализации)

    Для расчётных этапов (19+):
    - Возвращает расчётные данные
    - routes = список маршрутов с км
    """
    if stage_number < 1:
        raise HTTPException(400, "Номер этапа должен быть >= 1")

    details = get_stage_details(db, contract_number, stage_number)

    if not details:
        raise HTTPException(404, f"Этап {stage_number} не найден")

    return details


@router.get("/{contract_number}/historical")
async def get_historical_only(
    contract_number: str,
    db: Session = Depends(get_db)
):
    """
    Получить только исторические этапы (заблокированные).

    Этапы 1-18 (апрель 2022 - январь 2026).
    """
    # Передаём end_date до начала расчётов, чтобы получить только исторические
    end_dt = date(2026, 1, 31)
    summary = get_contract_summary(db, contract_number, end_dt)

    return {
        "contract_number": contract_number,
        "stages_count": len(summary.stages),
        "total_km": round(summary.historical_km, 2),
        "total_price": round(summary.historical_price, 2),
        "stages": [s.to_dict() for s in summary.stages],
    }


@router.get("/{contract_number}/calculated")
async def get_calculated_only(
    contract_number: str,
    end_date: str = Query(default=None, description="Конечная дата (YYYY-MM-DD)"),
    db: Session = Depends(get_db)
):
    """
    Получить только расчётные этапы.

    Этапы с 19 (февраль 2026 и далее).
    """
    end_dt = None
    if end_date:
        try:
            end_dt = date.fromisoformat(end_date)
        except ValueError:
            raise HTTPException(400, "Неверный формат даты")

    summary = get_contract_summary(db, contract_number, end_dt)

    # Фильтруем только расчётные
    calculated_stages = [s for s in summary.stages if not s.is_locked]

    return {
        "contract_number": contract_number,
        "stages_count": len(calculated_stages),
        "total_km": round(summary.calculated_km, 2),
        "total_price": round(summary.calculated_price, 2),
        "stages": [s.to_dict() for s in calculated_stages],
    }


@router.post("/{contract_number}/recalculate")
async def recalculate_stages(
    contract_number: str,
    end_date: str = Query(default=None, description="Конечная дата (YYYY-MM-DD)"),
    db: Session = Depends(get_db)
):
    """
    Принудительно пересчитать все этапы контракта.

    Рассчитывает этапы с февраля 2026 до end_date (по умолчанию июль 2028)
    и сохраняет результаты в БД.
    """
    contract = db.query(Contract).filter(Contract.number == contract_number).first()
    if not contract:
        raise HTTPException(404, "Контракт не найден")

    end_dt = None
    if end_date:
        try:
            end_dt = date.fromisoformat(end_date)
        except ValueError:
            raise HTTPException(400, "Неверный формат даты")

    stages_count = recalculate_contract_stages(db, contract_number, end_dt)

    return {
        "success": True,
        "contract_number": contract_number,
        "stages_recalculated": stages_count,
        "message": f"Пересчитано {stages_count} этапов",
    }


@router.get("/{contract_number}/status")
async def get_calculation_status(
    contract_number: str,
    db: Session = Depends(get_db)
):
    """
    Получить статус расчёта этапов.

    Показывает, есть ли сохранённые расчёты в БД.
    """
    contract = db.query(Contract).filter(Contract.number == contract_number).first()
    if not contract:
        raise HTTPException(404, "Контракт не найден")

    has_saved = has_calculated_stages(db, contract.id)

    return {
        "contract_number": contract_number,
        "has_calculated_stages": has_saved,
        "message": "Расчёты сохранены в БД" if has_saved else "Расчёты не выполнены, требуется применить ДС или выполнить ручной пересчёт",
    }


@router.post("/{contract_number}/stage/{stage_number}/close")
async def close_stage(
    contract_number: str,
    stage_number: int,
    request: StageCloseRequest,
    db: Session = Depends(get_db)
):
    """
    Закрыть этап (перевести в статус CLOSED).

    После закрытия этап не пересчитывается.
    Опционально можно передать финальные значения км и цены.
    """
    contract = db.query(Contract).filter(Contract.number == contract_number).first()
    if not contract:
        raise HTTPException(404, "Контракт не найден")

    stage = db.query(CalculatedStage).filter(
        CalculatedStage.contract_id == contract.id,
        CalculatedStage.stage == stage_number,
    ).first()

    if not stage:
        raise HTTPException(404, f"Этап {stage_number} не найден")

    if stage.status == StageStatus.CLOSED:
        raise HTTPException(400, "Этап уже закрыт")

    # Обновляем значения, если переданы
    if request.total_km is not None:
        stage.total_km = request.total_km
    if request.total_price is not None:
        stage.total_price = request.total_price

    # Закрываем этап
    stage.status = StageStatus.CLOSED

    db.commit()

    return {
        "success": True,
        "stage": stage_number,
        "total_km": stage.total_km,
        "total_price": stage.total_price,
        "status": "closed",
        "message": f"Этап {stage_number} закрыт",
    }


@router.post("/{contract_number}/stage/{stage_number}/reset")
async def reset_stage(
    contract_number: str,
    stage_number: int,
    db: Session = Depends(get_db),
):
    """
    Сбросить этап в расчётный статус (SAVED).

    Пересчитывает км и цену заново по текущим данным маршрутов.
    После сброса этап снова будет пересчитываться при изменении ДС.
    """
    if stage_number < 1:
        raise HTTPException(400, "Номер этапа должен быть >= 1")

    contract = db.query(Contract).filter(Contract.number == contract_number).first()
    if not contract:
        raise HTTPException(404, "Контракт не найден")

    updated = recalculate_single_stage(db, contract.id, stage_number)

    if not updated:
        raise HTTPException(404, f"Этап {stage_number} не найден в БД")

    return {
        "success": True,
        "stage": stage_number,
        "total_km": updated.total_km,
        "total_price": updated.total_price,
        "status": "saved",
        "message": f"Этап {stage_number} пересчитан и переведён в расчётный статус",
    }


