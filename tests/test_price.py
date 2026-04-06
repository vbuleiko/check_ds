"""Тесты для core/calculator/price.py."""
import pytest
from datetime import date
from decimal import Decimal

from core.calculator.price import (
    get_quarter_period,
    calculate_route_price,
    _parse_period_dates,
)


class TestGetQuarterPeriod:
    """Тесты для определения квартального периода."""

    def test_q1(self):
        assert get_quarter_period(date(2026, 1, 15)) == "01.01.2026-31.03.2026"
        assert get_quarter_period(date(2026, 3, 31)) == "01.01.2026-31.03.2026"

    def test_q2(self):
        assert get_quarter_period(date(2026, 4, 1)) == "01.04.2026-30.06.2026"
        assert get_quarter_period(date(2026, 6, 30)) == "01.04.2026-30.06.2026"

    def test_q3(self):
        assert get_quarter_period(date(2026, 7, 1)) == "01.07.2026-30.09.2026"

    def test_q4(self):
        assert get_quarter_period(date(2026, 10, 1)) == "01.10.2026-31.12.2026"
        assert get_quarter_period(date(2026, 12, 31)) == "01.10.2026-31.12.2026"


class TestCalculateRoutePrice:
    """Тесты для расчёта стоимости маршрута."""

    def test_basic_calculation(self):
        result = calculate_route_price(
            route="120",
            km=100.0,
            coefficient=Decimal("1.5"),
            capacity=50,
        )
        assert result == Decimal("7500.0")

    def test_zero_km(self):
        result = calculate_route_price("120", 0.0, Decimal("1.5"), 50)
        assert result == Decimal("0")

    def test_precision(self):
        result = calculate_route_price("120", 123.45, Decimal("2.345"), 75)
        expected = Decimal("123.45") * Decimal("75") * Decimal("2.345")
        assert result == expected


class TestParsePeriodDates:
    """Тесты для парсинга строки периода."""

    def test_standard_format(self):
        result = _parse_period_dates("01.01.2026-31.03.2026")
        assert result == (date(2026, 1, 1), date(2026, 3, 31))

    def test_with_dash(self):
        result = _parse_period_dates("01.04.2026–30.06.2026")
        assert result == (date(2026, 4, 1), date(2026, 6, 30))

    def test_with_em_dash(self):
        result = _parse_period_dates("01.07.2026—30.09.2026")
        assert result == (date(2026, 7, 1), date(2026, 9, 30))

    def test_too_short(self):
        assert _parse_period_dates("01.01.2026") is None

    def test_invalid_date(self):
        assert _parse_period_dates("32.13.2026-31.03.2026") is None
