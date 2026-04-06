"""
Сервис суммарных данных по контракту.

Объединяет исторические данные (константы) с расчётными (из БД или калькулятора).
"""
from datetime import date
from dataclasses import dataclass, field
from typing import Optional
from calendar import monthrange

from sqlalchemy.orm import Session

from core.historical_stages import (
    get_historical_stages,
    get_contract_config,
)
from db.models import Contract, CalculatedStage, StageStatus

import logging

logger = logging.getLogger(__name__)


@dataclass
class StageData:
    """Данные этапа (исторического или расчётного)."""
    stage: int
    year: int
    period: str
    date_from: date
    date_to: date
    max_km: float
    price: Optional[float] = None
    is_locked: bool = False
    source: str = "calculated"  # "historical" | "calculated" | "saved"
    status: str = "saved"  # "saved" | "not_closed" | "closed"

    def to_dict(self) -> dict:
        return {
            "stage": self.stage,
            "year": self.year,
            "period": self.period,
            "date_from": str(self.date_from),
            "date_to": str(self.date_to),
            "max_km": round(self.max_km, 2),
            "price": self.price,
            "is_locked": self.is_locked,
            "source": self.source,
            "status": self.status,
        }


@dataclass
class ContractSummary:
    """Суммарные данные по контракту."""
    contract_number: str
    stages: list[StageData] = field(default_factory=list)
    historical_km: float = 0.0
    calculated_km: float = 0.0
    historical_price: float = 0.0
    calculated_price: float = 0.0

    @property
    def total_km(self) -> float:
        return self.historical_km + self.calculated_km

    @property
    def total_price(self) -> float:
        return self.historical_price + self.calculated_price

    @property
    def stages_count(self) -> int:
        return len(self.stages)

    def to_dict(self) -> dict:
        return {
            "contract_number": self.contract_number,
            "stages_count": self.stages_count,
            "totals": {
                "historical_km": round(self.historical_km, 2),
                "calculated_km": round(self.calculated_km, 2),
                "total_km": round(self.total_km, 2),
                "historical_price": round(self.historical_price, 2),
                "calculated_price": round(self.calculated_price, 2),
                "total_price": round(self.total_price, 2),
            },
            "stages": [s.to_dict() for s in self.stages],
        }


MONTH_NAMES_RU = {
    1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
    5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
    9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь",
}


def _generate_monthly_periods(
    start_date: date,
    end_date: date
) -> list[tuple[int, int, date, date]]:
    """
    Генерирует список месячных периодов.

    Returns:
        Список кортежей (year, month, date_from, date_to)
    """
    periods = []
    current = date(start_date.year, start_date.month, 1)

    while current <= end_date:
        year = current.year
        month = current.month
        _, last_day = monthrange(year, month)

        d_from = date(year, month, 1)
        d_to = date(year, month, last_day)

        # Ограничиваем конец периода датой окончания
        if d_to > end_date:
            d_to = end_date

        periods.append((year, month, d_from, d_to))

        # Следующий месяц
        if month == 12:
            current = date(year + 1, 1, 1)
        else:
            current = date(year, month + 1, 1)

    return periods


def get_contract_summary(
    session: Session,
    contract_number: str,
    end_date: Optional[date] = None,
    force_recalculate: bool = False,
) -> ContractSummary:
    """
    Получить суммарные данные по контракту.

    Использует сохранённые в БД этапы. При первом обращении автоматически
    рассчитывает и сохраняет все этапы. Последующие загрузки берут данные
    из БД без пересчёта. Пересчёт запускается только кнопкой «Пересчитать».

    Args:
        session: Сессия БД
        contract_number: Номер контракта
        end_date: Конечная дата расчёта (по умолчанию 2028-07-31)
        force_recalculate: Не используется (оставлен для совместимости)

    Returns:
        ContractSummary с историческими и расчётными этапами
    """
    if end_date is None:
        end_date = date(2028, 7, 31)

    summary = ContractSummary(contract_number=contract_number)

    # 1. Загружаем исторические этапы
    historical = get_historical_stages(contract_number)
    for hs in historical:
        stage_data = StageData(
            stage=hs.stage,
            year=hs.year,
            period=hs.period,
            date_from=hs.date_from,
            date_to=hs.date_to,
            max_km=hs.max_km,
            price=hs.price,
            is_locked=True,
            source="historical",
            status="closed",  # Исторические этапы всегда закрыты
        )
        summary.stages.append(stage_data)
        summary.historical_km += hs.max_km
        if hs.price:
            summary.historical_price += hs.price

    # 2. Получаем контракт из БД
    contract = session.query(Contract).filter(
        Contract.number == contract_number
    ).first()

    if not contract:
        return summary

    # Получаем конфигурацию контракта
    config = get_contract_config(contract_number)
    calc_start = config["start_date"]

    if end_date < calc_start:
        return summary

    # 3. Загружаем сохранённые этапы из БД (индексируем по номеру)
    first_stage = config["first_stage"]

    saved_stages = session.query(CalculatedStage).filter(
        CalculatedStage.contract_id == contract.id,
        CalculatedStage.stage >= first_stage,
        CalculatedStage.date_to <= end_date,
    ).all()
    saved_by_stage = {cs.stage: cs for cs in saved_stages}

    # Если нет сохранённых этапов — автоматически рассчитываем и сохраняем
    if not saved_by_stage and not force_recalculate:
        from services.stages_calculator import recalculate_stages as _recalc
        logger.info(
            "Нет сохранённых этапов для контракта %s — автоматический расчёт",
            contract_number,
        )
        _recalc(session, contract.id, end_date)
        # Перезагружаем из БД
        saved_stages = session.query(CalculatedStage).filter(
            CalculatedStage.contract_id == contract.id,
            CalculatedStage.stage >= first_stage,
            CalculatedStage.date_to <= end_date,
        ).all()
        saved_by_stage = {cs.stage: cs for cs in saved_stages}

    # 4. Генерируем все периоды и заполняем из сохранённых данных
    monthly_periods = _generate_monthly_periods(calc_start, end_date)

    stage_num = first_stage

    for year, month, d_from, d_to in monthly_periods:
        if stage_num in saved_by_stage:
            # Используем сохранённые данные
            cs = saved_by_stage[stage_num]
            status = cs.status.value if cs.status else "saved"

            # Вычисляем total_km как сумму округлённых значений по маршрутам
            # (согласовано с логикой формирования столбца "Итого" в Excel)
            if cs.routes_data:
                total_km = sum(
                    round(data.get("total_km", 0), 2)
                    for data in cs.routes_data.values()
                )
            else:
                total_km = cs.total_km

            stage_data = StageData(
                stage=cs.stage,
                year=cs.year,
                period=cs.period_name,
                date_from=cs.date_from,
                date_to=cs.date_to,
                max_km=total_km,
                price=cs.total_price,
                is_locked=(status == "closed"),
                source="saved",
                status=status,
            )
            summary.stages.append(stage_data)
            summary.calculated_km += total_km
            if cs.total_price:
                summary.calculated_price += cs.total_price
        else:
            # Этап не найден в БД — показываем пустую строку-заглушку
            stage_data = StageData(
                stage=stage_num,
                year=year,
                period=MONTH_NAMES_RU[month],
                date_from=d_from,
                date_to=d_to,
                max_km=0,
                price=None,
                is_locked=False,
                source="not_calculated",
                status="not_calculated",
            )
            summary.stages.append(stage_data)

        stage_num += 1

    return summary


def get_stage_details(
    session: Session,
    contract_number: str,
    stage_number: int,
    force_recalculate: bool = False,
) -> Optional[dict]:
    """
    Получить детали конкретного этапа.

    Для исторических возвращает константы.
    Для расчётных — данные из БД. Если этап не рассчитан, возвращает заглушку.
    """
    # Исторический этап
    config = get_contract_config(contract_number)
    if stage_number < config["first_stage"]:
        historical = get_historical_stages(contract_number)
        for hs in historical:
            if hs.stage == stage_number:
                return {
                    "stage": hs.stage,
                    "year": hs.year,
                    "period": hs.period,
                    "date_from": str(hs.date_from),
                    "date_to": str(hs.date_to),
                    "max_km": round(hs.max_km, 2),
                    "price": hs.price,
                    "is_locked": True,
                    "source": "historical",
                    "routes": None,  # Нет детализации для исторических
                }
        return None

    # Расчётный этап — сначала пробуем получить из БД
    contract = session.query(Contract).filter(
        Contract.number == contract_number
    ).first()

    if not contract:
        return None

    # Получаем сохранённый этап из БД
    saved_stage = session.query(CalculatedStage).filter(
        CalculatedStage.contract_id == contract.id,
        CalculatedStage.stage == stage_number,
    ).first()

    if saved_stage:
        # Формируем routes из routes_data
        routes_data = []
        total_km = 0.0
        if saved_stage.routes_data:
            for route, data in sorted(saved_stage.routes_data.items()):
                route_km = round(data.get("total_km", 0), 2)
                total_km += route_km
                routes_data.append({
                    "route": route,
                    "total_km": route_km,
                    "forward_km": data.get("forward_km", 0),
                    "reverse_km": data.get("reverse_km", 0),
                    "total_trips": data.get("total_trips", 0),
                    "segments": data.get("segments", []),
                })
        else:
            total_km = saved_stage.total_km

        status = saved_stage.status.value if saved_stage.status else "saved"

        return {
            "stage": saved_stage.stage,
            "year": saved_stage.year,
            "period": saved_stage.period_name,
            "date_from": str(saved_stage.date_from),
            "date_to": str(saved_stage.date_to),
            "max_km": round(total_km, 2),
            "price": saved_stage.total_price,
            "is_locked": (status == "closed"),
            "source": "saved",
            "status": status,
            "calculated_at": str(saved_stage.calculated_at),
            "routes_count": len(routes_data),
            "routes": routes_data,
        }

    # Этап не рассчитан — возвращаем заглушку без расчёта на лету
    config = get_contract_config(contract_number)
    months_offset = stage_number - config["first_stage"]
    calc_start = config["start_date"]

    year = calc_start.year + (calc_start.month - 1 + months_offset) // 12
    month = (calc_start.month - 1 + months_offset) % 12 + 1

    return {
        "stage": stage_number,
        "year": year,
        "period": MONTH_NAMES_RU[month],
        "date_from": None,
        "date_to": None,
        "max_km": 0,
        "price": None,
        "is_locked": False,
        "source": "not_calculated",
        "status": "not_calculated",
        "routes_count": 0,
        "routes": [],
    }
