"""Тесты для core/calculator/kilometers.py."""
import pytest
from datetime import date

from core.calculator.kilometers import (
    DayCalculation,
    ParamsSegment,
    PeriodCalculation,
    get_season_for_date,
)
from db.models import SeasonType


class TestDayCalculation:
    """Тесты для дневного расчёта."""

    def test_total_trips(self):
        day = DayCalculation(date=date(2026, 3, 1), day_type=1, forward_trips=5, reverse_trips=3)
        assert day.total_trips == 8

    def test_total_km(self):
        day = DayCalculation(
            date=date(2026, 3, 1), day_type=1,
            forward_km=100.5, reverse_km=98.3
        )
        assert day.total_km == pytest.approx(198.8)

    def test_defaults(self):
        day = DayCalculation(date=date(2026, 3, 1), day_type=1)
        assert day.total_trips == 0
        assert day.total_km == 0.0


class TestParamsSegment:
    """Тесты для сегмента расчёта."""

    def test_total_km(self):
        seg = ParamsSegment(
            route_params_id=1,
            date_from=date(2026, 3, 1),
            date_to=date(2026, 3, 31),
            forward_km=500.0,
            reverse_km=480.0,
        )
        assert seg.total_km == pytest.approx(980.0)

    def test_to_dict(self):
        seg = ParamsSegment(
            route_params_id=1,
            date_from=date(2026, 3, 1),
            date_to=date(2026, 3, 31),
            source_appendix="Приложение №3 ДС53",
            length_forward=15.2,
            length_reverse=14.8,
            days_count=31,
            forward_trips=155,
            reverse_trips=155,
            forward_km=2356.0,
            reverse_km=2294.0,
        )
        d = seg.to_dict()
        assert d["route_params_id"] == 1
        assert d["source"] == "Приложение №3 ДС53"
        assert d["total_km"] == 4650.0


class TestPeriodCalculation:
    """Тесты для расчёта за период."""

    def test_empty_period(self):
        calc = PeriodCalculation(
            route="120",
            date_from=date(2026, 3, 1),
            date_to=date(2026, 3, 31),
        )
        assert calc.total_trips == 0
        assert calc.total_km == 0.0
        assert calc.total_forward_km == 0.0
        assert calc.total_reverse_km == 0.0

    def test_with_days(self):
        calc = PeriodCalculation(
            route="120",
            date_from=date(2026, 3, 1),
            date_to=date(2026, 3, 2),
        )
        calc.days = [
            DayCalculation(date=date(2026, 3, 1), day_type=1, forward_trips=5, reverse_trips=5, forward_km=100, reverse_km=95),
            DayCalculation(date=date(2026, 3, 2), day_type=1, forward_trips=5, reverse_trips=5, forward_km=100, reverse_km=95),
        ]
        assert calc.total_forward_trips == 10
        assert calc.total_reverse_trips == 10
        assert calc.total_trips == 20
        assert calc.total_forward_km == 200.0
        assert calc.total_reverse_km == 190.0
        assert calc.total_km == pytest.approx(390.0)

    def test_to_dict(self):
        calc = PeriodCalculation(
            route="120",
            date_from=date(2026, 3, 1),
            date_to=date(2026, 3, 1),
        )
        d = calc.to_dict()
        assert d["route"] == "120"
        assert d["total_km"] == 0.0


class TestGetSeasonForDate:
    """Тесты для определения сезона."""

    def test_winter_november(self):
        assert get_season_for_date(date(2026, 11, 16)) == SeasonType.WINTER

    def test_winter_january(self):
        assert get_season_for_date(date(2026, 1, 15)) == SeasonType.WINTER

    def test_summer_may(self):
        assert get_season_for_date(date(2026, 5, 1)) == SeasonType.SUMMER

    def test_summer_october(self):
        assert get_season_for_date(date(2026, 10, 1)) == SeasonType.SUMMER

    def test_boundary_winter_start(self):
        # 16 ноября — начало зимы
        assert get_season_for_date(date(2026, 11, 16)) == SeasonType.WINTER
        # 15 ноября — ещё лето
        assert get_season_for_date(date(2026, 11, 15)) == SeasonType.SUMMER

    def test_boundary_winter_end(self):
        # 14 апреля — ещё зима
        assert get_season_for_date(date(2026, 4, 14)) == SeasonType.WINTER
        # 15 апреля — начало лета
        assert get_season_for_date(date(2026, 4, 15)) == SeasonType.SUMMER

    def test_route_specific_season(self):
        # Маршрут 305 имеет другие сезонные периоды: зима 01.09-31.05, лето 01.06-31.08
        assert get_season_for_date(date(2026, 9, 1), route="305") == SeasonType.WINTER
        assert get_season_for_date(date(2026, 6, 15), route="305") == SeasonType.SUMMER

    def test_route_without_season_config(self):
        # Маршрут без сезонных настроек — используем общий календарь
        assert get_season_for_date(date(2026, 1, 15), route="999") == SeasonType.WINTER
