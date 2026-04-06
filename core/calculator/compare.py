"""
Модуль сравнения "было/стало".

Сравнивает объёмы транспортной работы до и после применения ДС.
"""
from datetime import date
from dataclasses import dataclass, field
from typing import Optional

from sqlalchemy.orm import Session

from db.models import Agreement, Contract, RouteParams, RouteTrips, SeasonType
from core.calculator.kilometers import (
    calculate_route_period,
    PeriodCalculation,
    get_route_params_for_date,
)
from core.calculator.calendar import get_dates_in_range
from core.constants import get_weekdays_for_day_type, parse_point_to_weekday


@dataclass
class RouteComparison:
    """Сравнение по одному маршруту."""
    route: str
    date_from: date
    date_to: date

    km_before: float = 0.0
    km_after: float = 0.0

    trips_before: int = 0
    trips_after: int = 0

    details_before: Optional[dict] = None
    details_after: Optional[dict] = None

    @property
    def km_diff(self) -> float:
        return self.km_after - self.km_before

    @property
    def km_diff_percent(self) -> float:
        if self.km_before == 0:
            return 0.0 if self.km_after == 0 else 100.0
        return round((self.km_diff / self.km_before) * 100, 2)

    @property
    def trips_diff(self) -> int:
        return self.trips_after - self.trips_before

    def to_dict(self) -> dict:
        return {
            "route": self.route,
            "date_from": str(self.date_from),
            "date_to": str(self.date_to),
            "km_before": round(self.km_before, 2),
            "km_after": round(self.km_after, 2),
            "km_diff": round(self.km_diff, 2),
            "km_diff_percent": self.km_diff_percent,
            "trips_before": self.trips_before,
            "trips_after": self.trips_after,
            "trips_diff": self.trips_diff,
        }


@dataclass
class AgreementComparison:
    """Сравнение по всему ДС."""
    agreement_id: int
    contract_number: str
    ds_number: str

    routes: list[RouteComparison] = field(default_factory=list)

    @property
    def total_km_before(self) -> float:
        # Округляем каждый маршрут до 2 знаков перед суммированием
        return sum(round(r.km_before, 2) for r in self.routes)

    @property
    def total_km_after(self) -> float:
        # Округляем каждый маршрут до 2 знаков перед суммированием
        return sum(round(r.km_after, 2) for r in self.routes)

    @property
    def total_km_diff(self) -> float:
        return self.total_km_after - self.total_km_before

    @property
    def total_trips_before(self) -> int:
        return sum(r.trips_before for r in self.routes)

    @property
    def total_trips_after(self) -> int:
        return sum(r.trips_after for r in self.routes)

    def to_dict(self) -> dict:
        return {
            "agreement_id": self.agreement_id,
            "contract_number": self.contract_number,
            "ds_number": self.ds_number,
            "total_km_before": round(self.total_km_before, 2),
            "total_km_after": round(self.total_km_after, 2),
            "total_km_diff": round(self.total_km_diff, 2),
            "total_trips_before": self.total_trips_before,
            "total_trips_after": self.total_trips_after,
            "routes": [r.to_dict() for r in self.routes],
        }


def compare_agreement(
    session: Session,
    agreement: Agreement,
    calculation_end_date: Optional[date] = None
) -> AgreementComparison:
    """
    Сравнивает объёмы до и после применения ДС.

    Args:
        session: Сессия БД
        agreement: ДС для сравнения
        calculation_end_date: Конец периода расчёта (по умолчанию — конец контракта)

    Returns:
        AgreementComparison с детализацией по маршрутам
    """
    contract = agreement.contract

    if calculation_end_date is None:
        calculation_end_date = contract.date_to or date(2028, 7, 14)

    result = AgreementComparison(
        agreement_id=agreement.id,
        contract_number=contract.number,
        ds_number=agreement.number,
    )

    json_data = agreement.json_data
    if not json_data:
        return result

    # Собираем изменения с деньгами (влияют на объём)
    changes_with_money = json_data.get("change_with_money", [])
    changes_no_appendix = json_data.get("change_with_money_no_appendix", [])
    appendices = json_data.get("appendices", {})

    # Обрабатываем изменения с приложениями
    for change in changes_with_money:
        appendix_id = change.get("appendix")
        route = change.get("route")

        if not route:
            continue

        # Определяем период действия
        date_from = _parse_date(change.get("date_from"))
        date_to = _parse_date(change.get("date_to"))
        date_on = _parse_date(change.get("date_on"))

        if date_on:
            # Изменение на один день
            period_start = date_on
            period_end = date_on
        elif date_from:
            period_start = date_from
            period_end = date_to or calculation_end_date
        else:
            continue

        # Рассчитываем "было"
        calc_before = calculate_route_period(
            session, contract.id, route, period_start, period_end
        )

        # Для "стало" нужно применить новые параметры из приложения
        # Это сложнее — нужно временно создать параметры и пересчитать
        # Пока упрощённо: берём данные из appendix
        appendix_data = appendices.get(appendix_id, {})
        km_after = _estimate_km_from_appendix(
            appendix_data, period_start, period_end
        )

        comparison = RouteComparison(
            route=route,
            date_from=period_start,
            date_to=period_end,
            km_before=calc_before.total_km,
            km_after=km_after,
            trips_before=calc_before.total_trips,
            trips_after=0,  # TODO: рассчитать
            details_before=calc_before.to_dict(),
        )
        result.routes.append(comparison)

    # Обрабатываем изменения без приложений (праздники)
    for change in changes_no_appendix:
        route = change.get("route")
        point = change.get("point", "")

        if not route:
            continue

        date_from = _parse_date(change.get("date_from"))
        date_to = _parse_date(change.get("date_to"))
        date_on = _parse_date(change.get("date_on"))

        if date_on:
            period_start = date_on
            period_end = date_on
        elif date_from:
            period_start = date_from
            period_end = date_to or date_from
        else:
            continue

        # Рассчитываем "было" (обычный день)
        calc_before = calculate_route_period(
            session, contract.id, route, period_start, period_end
        )

        # "Стало" — работа по другому графику
        # Нужно пересчитать с изменённым типом дня
        new_day_type = parse_point_to_weekday(point)
        if new_day_type:
            # TODO: пересчитать с новым типом дня
            km_after = calc_before.total_km  # Пока заглушка
        else:
            km_after = calc_before.total_km

        comparison = RouteComparison(
            route=route,
            date_from=period_start,
            date_to=period_end,
            km_before=calc_before.total_km,
            km_after=km_after,
            trips_before=calc_before.total_trips,
        )
        result.routes.append(comparison)

    return result


def _parse_date(value) -> Optional[date]:
    """Парсит дату из строки."""
    if not value:
        return None
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            # YYYY-MM-DD
            parts = value.split("-")
            if len(parts) == 3:
                return date(int(parts[0]), int(parts[1]), int(parts[2]))
        except (ValueError, IndexError):
            pass
    return None


def _estimate_km_from_appendix(
    appendix_data: dict,
    date_from: date,
    date_to: date
) -> float:
    """
    Оценивает км из данных приложения.

    Упрощённый расчёт: берём средние рейсы и умножаем на дни.
    """
    if not appendix_data:
        return 0.0

    length_forward = appendix_data.get("length_forward", 0) or 0
    length_reverse = appendix_data.get("length_reverse", 0) or 0
    num_of_types = appendix_data.get("num_of_types", 0)

    # Считаем дни
    dates = get_dates_in_range(date_from, date_to)
    workdays = sum(1 for d in dates if d.isoweekday() <= 5)
    weekends = len(dates) - workdays

    total_km = 0.0

    for type_num in range(1, num_of_types + 1):
        type_name = appendix_data.get(f"type_{type_num}_name", "").lower()

        forward = appendix_data.get(f"type_{type_num}_forward_number", 0) or 0
        reverse = appendix_data.get(f"type_{type_num}_reverse_number", 0) or 0

        # Определяем, сколько дней этого типа
        if "рабочие" in type_name and "выходные" not in type_name:
            days = workdays
        elif "выходные" in type_name or "воскресн" in type_name:
            days = weekends
        else:
            days = len(dates)

        km = (forward * length_forward + reverse * length_reverse) * days
        total_km += km

    return total_km
