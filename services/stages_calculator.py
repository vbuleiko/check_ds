"""
Сервис расчёта и сохранения этапов контракта.

Рассчитывает этапы с февраля 2026 и сохраняет их в БД.
Вызывается при применении ДС или вручную.
"""
from datetime import date, datetime, timezone
from calendar import monthrange
from typing import Optional

from sqlalchemy.orm import Session

from db.models import Contract, CalculatedStage, Agreement, StageStatus
from core.historical_stages import CALCULATION_START_DATE, LAST_HISTORICAL_STAGE, get_contract_config
from core.calculator.kilometers import calculate_contract_period
from core.calculator.price import calculate_stage_price, preload_price_data


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

        if d_to > end_date:
            d_to = end_date

        periods.append((year, month, d_from, d_to))

        if month == 12:
            current = date(year + 1, 1, 1)
        else:
            current = date(year, month + 1, 1)

    return periods


def recalculate_stages(
    session: Session,
    contract_id: int,
    end_date: Optional[date] = None,
    source_agreement_id: Optional[int] = None,
) -> int:
    """
    Пересчитывает все этапы контракта и сохраняет в БД.

    Args:
        session: Сессия БД
        contract_id: ID контракта
        end_date: Конечная дата расчёта (по умолчанию 2028-07-31)
        source_agreement_id: ID ДС, вызвавшего пересчёт

    Returns:
        Количество рассчитанных/обновлённых этапов
    """
    if end_date is None:
        end_date = date(2028, 7, 31)

    contract = session.query(Contract).filter(Contract.id == contract_id).first()
    if not contract:
        return 0

    # Получаем конфигурацию контракта
    config = get_contract_config(contract.number)
    calc_start = config["start_date"]
    first_stage = config["first_stage"]

    if end_date < calc_start:
        return 0

    monthly_periods = _generate_monthly_periods(calc_start, end_date)

    # Предзагружаем данные для расчёта цены
    preload_price_data(contract.number)

    stage_num = first_stage
    updated_count = 0

    for year, month, d_from, d_to in monthly_periods:
        # Рассчитываем км за месяц
        results = calculate_contract_period(session, contract_id, d_from, d_to)

        # Округляем каждый маршрут до 2 знаков перед суммированием
        total_km = sum(round(r.total_km, 2) for r in results.values())

        # Рассчитываем цену
        route_km = {route: calc.total_km for route, calc in results.items()}
        stage_price = calculate_stage_price(route_km, d_to, contract.number)

        # Формируем детализацию по маршрутам
        routes_data = {}
        for route, calc in sorted(results.items()):
            routes_data[route] = {
                "total_km": round(calc.total_km, 2),
                "forward_km": round(calc.total_forward_km, 2),
                "reverse_km": round(calc.total_reverse_km, 2),
                "total_trips": calc.total_trips,
                "forward_trips": calc.total_forward_trips,
                "reverse_trips": calc.total_reverse_trips,
                "segments": [s.to_dict() for s in calc.segments] if calc.segments else [],
            }

        # Ищем существующий этап или создаём новый
        existing = session.query(CalculatedStage).filter(
            CalculatedStage.contract_id == contract_id,
            CalculatedStage.stage == stage_num,
        ).first()

        if existing:
            # Не пересчитываем закрытые этапы
            if existing.status == StageStatus.CLOSED:
                stage_num += 1
                continue

            # Обновляем
            existing.total_km = total_km
            existing.total_price = round(stage_price, 2) if stage_price else None
            existing.routes_data = routes_data
            existing.calculated_at = datetime.now(timezone.utc)
            existing.source_agreement_id = source_agreement_id
        else:
            # Создаём новый (статус SAVED по умолчанию)
            new_stage = CalculatedStage(
                contract_id=contract_id,
                stage=stage_num,
                year=year,
                month=month,
                period_name=MONTH_NAMES_RU[month],
                date_from=d_from,
                date_to=d_to,
                total_km=total_km,
                total_price=round(stage_price, 2) if stage_price else None,
                routes_data=routes_data,
                source_agreement_id=source_agreement_id,
            )
            session.add(new_stage)

        updated_count += 1
        stage_num += 1

    session.commit()
    return updated_count


def recalculate_contract_stages(
    session: Session,
    contract_number: str,
    end_date: Optional[date] = None,
    source_agreement_id: Optional[int] = None,
) -> int:
    """
    Пересчитывает этапы по номеру контракта.

    Args:
        session: Сессия БД
        contract_number: Номер контракта (222, 219 и т.д.)
        end_date: Конечная дата расчёта
        source_agreement_id: ID ДС, вызвавшего пересчёт

    Returns:
        Количество рассчитанных/обновлённых этапов
    """
    contract = session.query(Contract).filter(
        Contract.number == contract_number
    ).first()

    if not contract:
        return 0

    return recalculate_stages(
        session,
        contract.id,
        end_date,
        source_agreement_id,
    )


def recalculate_single_stage(
    session: Session,
    contract_id: int,
    stage_number: int,
) -> Optional["CalculatedStage"]:
    """
    Пересчитывает один этап контракта и устанавливает статус SAVED.

    Args:
        session: Сессия БД
        contract_id: ID контракта
        stage_number: Номер этапа

    Returns:
        Обновлённый CalculatedStage или None если этап не найден
    """
    existing = session.query(CalculatedStage).filter(
        CalculatedStage.contract_id == contract_id,
        CalculatedStage.stage == stage_number,
    ).first()

    if not existing:
        return None

    contract = session.query(Contract).filter(Contract.id == contract_id).first()
    if not contract:
        return None

    preload_price_data(contract.number)

    d_from = existing.date_from
    d_to = existing.date_to

    results = calculate_contract_period(session, contract_id, d_from, d_to)

    total_km = sum(round(r.total_km, 2) for r in results.values())

    route_km = {route: calc.total_km for route, calc in results.items()}
    stage_price = calculate_stage_price(route_km, d_to, contract.number)

    routes_data = {}
    for route, calc in sorted(results.items()):
        routes_data[route] = {
            "total_km": round(calc.total_km, 2),
            "forward_km": round(calc.total_forward_km, 2),
            "reverse_km": round(calc.total_reverse_km, 2),
            "total_trips": calc.total_trips,
            "forward_trips": calc.total_forward_trips,
            "reverse_trips": calc.total_reverse_trips,
            "segments": [s.to_dict() for s in calc.segments] if calc.segments else [],
        }

    existing.total_km = total_km
    existing.total_price = round(stage_price, 2) if stage_price else None
    existing.routes_data = routes_data
    existing.status = StageStatus.SAVED
    existing.calculated_at = datetime.now(timezone.utc)

    session.commit()
    return existing


def get_calculated_stages(
    session: Session,
    contract_id: int,
) -> list[CalculatedStage]:
    """
    Получает все рассчитанные этапы контракта из БД.

    Args:
        session: Сессия БД
        contract_id: ID контракта

    Returns:
        Список рассчитанных этапов
    """
    return session.query(CalculatedStage).filter(
        CalculatedStage.contract_id == contract_id
    ).order_by(CalculatedStage.stage).all()


def has_calculated_stages(session: Session, contract_id: int) -> bool:
    """Проверяет, есть ли рассчитанные этапы для контракта."""
    return session.query(CalculatedStage).filter(
        CalculatedStage.contract_id == contract_id
    ).first() is not None
