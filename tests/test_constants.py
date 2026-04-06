"""Тесты для core/constants.py."""
import pytest
from core.constants import (
    get_weekdays_for_day_type,
    get_weekdays_for_type_extended,
    parse_point_to_weekday,
)


class TestGetWeekdaysForDayType:
    """Тесты для простого маппинга типов дней."""

    def test_exact_match(self):
        assert get_weekdays_for_day_type("Рабочие дни") == [1, 2, 3, 4, 5]

    def test_exact_match_friday(self):
        assert get_weekdays_for_day_type("Пятница") == [5]

    def test_exact_match_weekends(self):
        assert get_weekdays_for_day_type("Выходные дни") == [6, 7]

    def test_fuzzy_match_case_insensitive(self):
        assert get_weekdays_for_day_type("рабочие дни") == [1, 2, 3, 4, 5]

    def test_fuzzy_match_partial(self):
        assert get_weekdays_for_day_type("Субботние дни (и ещё текст)") == [6]

    def test_unknown_returns_all_days(self):
        assert get_weekdays_for_day_type("неизвестный тип") == [1, 2, 3, 4, 5, 6, 7]


class TestGetWeekdaysForTypeExtended:
    """Тесты для расширенного маппинга (парсинг документов)."""

    def test_workdays(self):
        assert get_weekdays_for_type_extended("Рабочие дни") == [1, 2, 3, 4, 5]

    def test_workdays_lowercase(self):
        assert get_weekdays_for_type_extended("рабочие дни") == [1, 2, 3, 4, 5]

    def test_friday(self):
        assert get_weekdays_for_type_extended("Пятница") == [5]

    def test_saturday(self):
        assert get_weekdays_for_type_extended("Суббота") == [6]

    def test_sunday(self):
        assert get_weekdays_for_type_extended("Воскресенье") == [7]

    def test_weekends(self):
        assert get_weekdays_for_type_extended("Выходные дни") == [6, 7]

    def test_daily(self):
        assert get_weekdays_for_type_extended("Ежедневно") == [1, 2, 3, 4, 5, 6, 7]

    def test_workdays_and_saturdays(self):
        assert get_weekdays_for_type_extended("Рабочие и субботние дни") == [1, 2, 3, 4, 5, 6]

    def test_workdays_except_friday(self):
        assert get_weekdays_for_type_extended("Рабочие дни кроме пятницы") == [1, 2, 3, 4]

    def test_all_days(self):
        assert get_weekdays_for_type_extended("Рабочие, выходные и праздничные дни") == [1, 2, 3, 4, 5, 6, 7]

    def test_abbreviation_pn(self):
        assert get_weekdays_for_type_extended("Пн") == [1]

    def test_abbreviation_sb(self):
        assert get_weekdays_for_type_extended("Сб") == [6]

    def test_abbreviation_vs(self):
        assert get_weekdays_for_type_extended("Вс") == [7]

    def test_krome_pattern_generic(self):
        """Тест на паттерн 'кроме' для произвольных типов."""
        result = get_weekdays_for_type_extended("Рабочие дни кроме среды")
        assert 3 not in result
        assert result == [1, 2, 4, 5]

    def test_unknown_returns_empty(self):
        assert get_weekdays_for_type_extended("абракадабра") == []

    def test_whitespace_handling(self):
        assert get_weekdays_for_type_extended("  Рабочие дни  ") == [1, 2, 3, 4, 5]


class TestParsePointToWeekday:
    """Тесты для parse_point_to_weekday."""

    def test_friday_reference(self):
        text = "по графику движения соответствующему графику движения пятниц"
        assert parse_point_to_weekday(text) == 5

    def test_saturday_reference(self):
        text = "по графику движения субботних дней"
        assert parse_point_to_weekday(text) == 6

    def test_sunday_reference(self):
        text = "считать воскресным днём"
        assert parse_point_to_weekday(text) == 7

    def test_no_match(self):
        assert parse_point_to_weekday("обычный текст") is None
