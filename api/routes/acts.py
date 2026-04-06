"""API для загрузки и обработки актов приёмки выполненных работ."""
import tempfile
from calendar import monthrange
from datetime import date
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, UploadFile, File, Depends, HTTPException
from pydantic import BaseModel
from sqlalchemy.orm import Session

from db.database import get_db
from db.models import Contract, CalculatedStage, StageStatus
from core.parser.act_docx import parse_act_docx
from core.historical_stages import get_historical_stages, LAST_HISTORICAL_STAGE, get_contract_config
from services.stages_calculator import recalculate_stages

router = APIRouter()


def _find_stage_for_period(
    db: Session,
    contract_id: int,
    contract_number: str,
    year: int,
    month: int,
) -> Optional[dict]:
    """Находит данные этапа для заданного контракта и месяца/года."""
    _, last_day = monthrange(year, month)
    period_from = date(year, month, 1)
    period_to = date(year, month, last_day)

    # 1. Исторические этапы (до Jan 2026)
    for hs in get_historical_stages(contract_number):
        if hs.date_from <= period_from and period_to <= hs.date_to:
            return {
                'stage': hs.stage,
                'period': hs.period,
                'year': hs.year,
                'date_from': str(hs.date_from),
                'date_to': str(hs.date_to),
                'calculated_km': round(hs.max_km, 2),
                'calculated_price': hs.price,
                'is_locked': True,
                'status': 'closed',
            }

    # 2. Расчётные этапы (с даты начала расчётов)
    saved = db.query(CalculatedStage).filter(
        CalculatedStage.contract_id == contract_id,
        CalculatedStage.year == year,
        CalculatedStage.month == month,
    ).first()

    # Если этап не найден — пробуем рассчитать автоматически
    if not saved:
        config = get_contract_config(contract_number)
        if period_from >= config.get("start_date", date(2099, 1, 1)):
            recalculate_stages(db, contract_id, end_date=period_to)
            saved = db.query(CalculatedStage).filter(
                CalculatedStage.contract_id == contract_id,
                CalculatedStage.year == year,
                CalculatedStage.month == month,
            ).first()

    if saved:
        if saved.routes_data:
            total_km = round(
                sum(round(d.get('total_km', 0), 2) for d in saved.routes_data.values()), 2
            )
        else:
            total_km = round(saved.total_km, 2)
        return {
            'stage': saved.stage,
            'period': saved.period_name,
            'year': saved.year,
            'date_from': str(saved.date_from),
            'date_to': str(saved.date_to),
            'calculated_km': total_km,
            'calculated_price': saved.total_price,
            'is_locked': saved.status == StageStatus.CLOSED,
            'status': saved.status.value if saved.status else 'saved',
        }

    return None


@router.post("/parse")
async def parse_acts(
    files: list[UploadFile] = File(...),
    db: Session = Depends(get_db),
):
    """
    Принимает несколько .docx актов, парсит их и возвращает данные
    для сравнения с расчётными значениями этапов.
    """
    results = []

    for upload_file in files:
        filename = upload_file.filename or 'unknown.docx'

        if not filename.lower().endswith('.docx'):
            results.append({
                'filename': filename,
                'error': 'Поддерживаются только .docx файлы',
            })
            continue

        content = await upload_file.read()

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir) / filename
            tmp_path.write_bytes(content)

            try:
                act = parse_act_docx(tmp_path)
            except Exception as e:
                results.append({
                    'filename': filename,
                    'error': f'Ошибка парсинга: {e}',
                })
                continue

        if act.parse_errors:
            results.append({
                'filename': act.filename,
                'contract_number': act.contract_number or None,
                'error': '; '.join(act.parse_errors),
                'act_data': {
                    'contract_full_number': act.contract_full_number,
                    'period_year': act.period_year,
                    'period_month': act.period_month,
                    'period_name': act.period_name,
                    'total_km_fact': act.total_km_fact,
                    'total_price': act.total_price,
                },
            })
            continue

        contract = db.query(Contract).filter(
            Contract.number == act.contract_number
        ).first()

        stage_data = None
        if contract and act.period_year and act.period_month:
            stage_data = _find_stage_for_period(
                db, contract.id, act.contract_number,
                act.period_year, act.period_month,
            )

        results.append({
            'filename': act.filename,
            'contract_number': act.contract_number,
            'contract_found': contract is not None,
            'act_data': {
                'contract_full_number': act.contract_full_number,
                'period_year': act.period_year,
                'period_month': act.period_month,
                'period_name': act.period_name,
                'total_km_fact': act.total_km_fact,
                'total_price': act.total_price,
            },
            'stage': stage_data,
        })

    return results


class ConfirmItem(BaseModel):
    contract_number: str
    stage_number: int
    total_km: Optional[float] = None
    total_price: Optional[float] = None


@router.post("/confirm")
async def confirm_acts(
    items: list[ConfirmItem],
    db: Session = Depends(get_db),
):
    """
    Подтверждает данные из актов: обновляет значения км/цены и закрывает этапы.
    Работает только для расчётных этапов (не для исторических).
    """
    results = []

    for item in items:
        contract = db.query(Contract).filter(
            Contract.number == item.contract_number
        ).first()

        if not contract:
            results.append({
                'contract_number': item.contract_number,
                'stage': item.stage_number,
                'success': False,
                'error': 'Контракт не найден',
            })
            continue

        stage = db.query(CalculatedStage).filter(
            CalculatedStage.contract_id == contract.id,
            CalculatedStage.stage == item.stage_number,
        ).first()

        if not stage:
            results.append({
                'contract_number': item.contract_number,
                'stage': item.stage_number,
                'success': False,
                'error': f'Этап {item.stage_number} не найден '
                         f'(исторические этапы не подлежат изменению)',
            })
            continue

        if item.total_price is not None:
            stage.total_price = item.total_price
        stage.status = StageStatus.CLOSED
        db.commit()

        results.append({
            'contract_number': item.contract_number,
            'stage': item.stage_number,
            'success': True,
            'total_km': stage.total_km,
            'total_price': stage.total_price,
        })

    return {'results': results}


@router.get("/history")
async def acts_history(db: Session = Depends(get_db)):
    """Возвращает все закрытые (подтверждённые) этапы с данными из актов."""
    stages = (
        db.query(CalculatedStage)
        .filter(CalculatedStage.status == StageStatus.CLOSED)
        .order_by(CalculatedStage.contract_id, CalculatedStage.year, CalculatedStage.month)
        .all()
    )

    result = []
    for s in stages:
        contract = db.query(Contract).filter(Contract.id == s.contract_id).first()
        if s.routes_data:
            total_km = round(
                sum(round(d.get('total_km', 0), 2) for d in s.routes_data.values()), 2
            )
        else:
            total_km = round(s.total_km, 2) if s.total_km else 0.0
        result.append({
            'contract_number': contract.number if contract else '?',
            'stage': s.stage,
            'period_name': s.period_name,
            'year': s.year,
            'date_from': str(s.date_from),
            'date_to': str(s.date_to),
            'calculated_km': total_km,
            'total_price': s.total_price,
        })

    return result
