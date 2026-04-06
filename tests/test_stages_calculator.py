"""Тесты для services/stages_calculator.py."""
import pytest
from datetime import date

from services.stages_calculator import _generate_monthly_periods


class TestGenerateMonthlyPeriods:
    def test_single_month(self):
        periods = _generate_monthly_periods(date(2026, 1, 1), date(2026, 1, 31))
        assert len(periods) == 1
        year, month, d_from, d_to = periods[0]
        assert year == 2026
        assert month == 1
        assert d_from == date(2026, 1, 1)
        assert d_to == date(2026, 1, 31)

    def test_three_months(self):
        periods = _generate_monthly_periods(date(2026, 1, 1), date(2026, 3, 31))
        assert len(periods) == 3
        assert periods[0][1] == 1
        assert periods[1][1] == 2
        assert periods[2][1] == 3

    def test_year_boundary(self):
        """Декабрь → Январь следующего года."""
        periods = _generate_monthly_periods(date(2026, 11, 1), date(2027, 2, 28))
        months = [(p[0], p[1]) for p in periods]
        assert (2026, 11) in months
        assert (2026, 12) in months
        assert (2027, 1) in months
        assert (2027, 2) in months

    def test_partial_first_month(self):
        """Начало с середины месяца — d_from = start_date, d_to = конец месяца."""
        periods = _generate_monthly_periods(date(2026, 2, 15), date(2026, 3, 31))
        assert len(periods) == 2
        # Первый период — полный февраль начиная с 1-го (алгоритм берёт 1-е число месяца start)
        assert periods[0][2] == date(2026, 2, 1)

    def test_end_truncates_last_month(self):
        """Если end_date < конца последнего месяца — d_to обрезается."""
        periods = _generate_monthly_periods(date(2026, 3, 1), date(2026, 3, 15))
        assert len(periods) == 1
        assert periods[0][3] == date(2026, 3, 15)

    def test_start_equals_end_single_period(self):
        periods = _generate_monthly_periods(date(2026, 6, 15), date(2026, 6, 15))
        assert len(periods) == 1
        assert periods[0][3] == date(2026, 6, 15)

    def test_february_leap_year(self):
        periods = _generate_monthly_periods(date(2028, 2, 1), date(2028, 2, 29))
        assert len(periods) == 1
        assert periods[0][3] == date(2028, 2, 29)

    def test_february_non_leap_year(self):
        periods = _generate_monthly_periods(date(2027, 2, 1), date(2027, 2, 28))
        assert len(periods) == 1
        assert periods[0][3] == date(2027, 2, 28)

    def test_period_tuple_structure(self):
        periods = _generate_monthly_periods(date(2026, 5, 1), date(2026, 5, 31))
        year, month, d_from, d_to = periods[0]
        assert isinstance(year, int)
        assert isinstance(month, int)
        assert isinstance(d_from, date)
        assert isinstance(d_to, date)
