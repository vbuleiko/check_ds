"""Тесты для core/checker/internal.py."""
import pytest

from core.checker.internal import (
    CheckItem,
    CheckResult,
    check_sums_consistency,
    check_probeg_consistency,
    check_appendix_arithmetic,
    check_appendix_numbering,
    check_ds_numbers_in_appendices,
    check_json,
)


class TestCheckItem:
    def test_ok_when_no_errors(self):
        item = CheckItem(id="x", label="x")
        assert item.ok is True

    def test_not_ok_when_has_errors(self):
        item = CheckItem(id="x", label="x", errors=["err"])
        assert item.ok is False

    def test_has_warnings(self):
        item = CheckItem(id="x", label="x", warnings=["w"])
        assert item.has_warnings is True

    def test_to_dict_structure(self):
        item = CheckItem(id="test", label="Test", errors=["e"], warnings=["w"], detail="d")
        d = item.to_dict()
        assert d["id"] == "test"
        assert d["label"] == "Test"
        assert d["ok"] is False
        assert d["has_warnings"] is True
        assert d["errors"] == ["e"]
        assert d["warnings"] == ["w"]
        assert d["detail"] == "d"


class TestCheckResult:
    def test_is_valid_when_no_errors(self):
        result = CheckResult()
        assert result.is_valid is True

    def test_not_valid_when_has_errors(self):
        result = CheckResult(errors=["err"])
        assert result.is_valid is False

    def test_add_check_accumulates_errors_and_warnings(self):
        result = CheckResult()
        item = CheckItem(id="x", label="x", errors=["e1"], warnings=["w1"])
        result.add_check(item)
        assert "e1" in result.errors
        assert "w1" in result.warnings
        assert item in result.checks

    def test_add_check_ok_item_does_not_add_errors(self):
        result = CheckResult()
        item = CheckItem(id="x", label="x")
        result.add_check(item)
        assert result.errors == []
        assert result.is_valid is True


class TestCheckSumsConsistency:
    def test_all_equal_sums_ok(self):
        data = {"general": {"sum_text": 100.0, "sum_finansirovanie_text": 100.0, "sum_etapy": 100.0}}
        item = check_sums_consistency(data)
        assert item.ok

    def test_different_sums_error(self):
        data = {"general": {"sum_text": 100.0, "sum_etapy": 200.0}}
        item = check_sums_consistency(data)
        assert not item.ok
        assert len(item.errors) == 1

    def test_no_sum_fields_no_data(self):
        data = {"general": {}}
        item = check_sums_consistency(data)
        assert item.ok
        assert item.detail == "Нет данных"

    def test_single_sum_field_ok(self):
        data = {"general": {"sum_text": 500.0}}
        item = check_sums_consistency(data)
        assert item.ok

    def test_empty_data(self):
        item = check_sums_consistency({})
        assert item.ok


class TestCheckProbegConsistency:
    def test_equal_probegs_ok(self):
        data = {"general": {"probeg_sravnenie": 1000.0, "probeg_etapy": 1000.0}}
        item = check_probeg_consistency(data)
        assert item.ok

    def test_different_probegs_error(self):
        data = {"general": {"probeg_sravnenie": 1000.0, "probeg_etapy": 999.0}}
        item = check_probeg_consistency(data)
        assert not item.ok

    def test_no_probeg_fields(self):
        data = {"general": {}}
        item = check_probeg_consistency(data)
        assert item.ok
        assert item.detail == "Нет данных"


class TestCheckAppendixArithmetic:
    def _make_data(self, forward_number, reverse_number, sum_number,
                   forward_probeg, reverse_probeg, sum_probeg,
                   length_forward=None, length_reverse=None, length_sum=None):
        appendix = {
            "appendix_num": "1",
            "num_of_types": 1,
            "type_1_name": "Рабочие дни",
            "type_1_forward_number": forward_number,
            "type_1_reverse_number": reverse_number,
            "type_1_sum_number": sum_number,
            "type_1_forward_probeg": forward_probeg,
            "type_1_reverse_probeg": reverse_probeg,
            "type_1_sum_probeg": sum_probeg,
        }
        if length_forward is not None:
            appendix["length_forward"] = length_forward
        if length_reverse is not None:
            appendix["length_reverse"] = length_reverse
        if length_sum is not None:
            appendix["length_sum"] = length_sum
        return {"appendices": {"1": appendix}}

    def test_correct_arithmetic_ok(self):
        data = self._make_data(10, 5, 15, 100.0, 50.0, 150.0, length_forward=10.0, length_reverse=10.0, length_sum=20.0)
        item = check_appendix_arithmetic(data)
        assert item.ok

    def test_wrong_sum_number_error(self):
        data = self._make_data(10, 5, 20, 100.0, 50.0, 150.0)
        item = check_appendix_arithmetic(data)
        assert not item.ok
        assert any("Кол-во рейсов" in e for e in item.errors)

    def test_wrong_sum_probeg_error(self):
        data = self._make_data(10, 5, 15, 100.0, 50.0, 200.0)
        item = check_appendix_arithmetic(data)
        assert not item.ok
        assert any("Пробег общий" in e for e in item.errors)

    def test_wrong_length_calculation_error(self):
        # length_forward=10, forward_number=10 → expected probeg=100, but given 50
        data = self._make_data(10, 0, 10, 50.0, 0.0, 50.0, length_forward=10.0)
        item = check_appendix_arithmetic(data)
        assert not item.ok
        assert any("Пробег от НП" in e for e in item.errors)

    def test_no_appendices_ok(self):
        item = check_appendix_arithmetic({})
        assert item.ok

    def test_wrong_length_sum_error(self):
        data = self._make_data(0, 0, 0, 0.0, 0.0, 0.0, length_forward=10.0, length_reverse=5.0, length_sum=20.0)
        item = check_appendix_arithmetic(data)
        assert not item.ok
        assert any("Протяжённость всего" in e for e in item.errors)


class TestCheckAppendixNumbering:
    def test_sequential_numbering_ok(self):
        data = {"appendices": {"1": {}, "2": {}, "3": {}}}
        item = check_appendix_numbering(data)
        assert item.ok
        assert "3" in item.detail

    def test_missing_appendix_error(self):
        data = {"appendices": {"1": {}, "3": {}}}
        item = check_appendix_numbering(data)
        assert not item.ok
        assert any("2" in e for e in item.errors)

    def test_empty_appendices_ok(self):
        item = check_appendix_numbering({"appendices": {}})
        assert item.ok

    def test_no_appendices_key_ok(self):
        item = check_appendix_numbering({})
        assert item.ok

    def test_km_contract_missing_appendix_ok(self):
        """Для ГК219/220/222 отсутствие приложения-Excel (km_data) допустимо."""
        data = {
            "appendices": {"1": {}, "3": {}},
            "general": {"contract_number": "2190001"},
            "km_data": {"appendix_number": "2"},
        }
        item = check_appendix_numbering(data)
        assert item.ok


class TestCheckDsNumbersInAppendices:
    def test_matching_ds_numbers_ok(self):
        data = {
            "general": {"ds_number": "5"},
            "appendices": {
                "1": {"appendix_num": "1", "ds_num": "5"},
                "2": {"appendix_num": "2", "ds_num": "5"},
            }
        }
        item = check_ds_numbers_in_appendices(data)
        assert item.ok

    def test_mismatched_ds_number_error(self):
        data = {
            "general": {"ds_number": "5"},
            "appendices": {
                "1": {"appendix_num": "1", "ds_num": "6"},
            }
        }
        item = check_ds_numbers_in_appendices(data)
        assert not item.ok
        assert len(item.errors) == 1

    def test_missing_ds_num_field_skipped(self):
        """Если поле ds_num отсутствует — пропускаем приложение."""
        data = {
            "general": {"ds_number": "5"},
            "appendices": {"1": {"appendix_num": "1"}},
        }
        item = check_ds_numbers_in_appendices(data)
        assert item.ok

    def test_no_ds_number_in_general(self):
        data = {
            "general": {},
            "appendices": {"1": {"ds_num": "5"}},
        }
        item = check_ds_numbers_in_appendices(data)
        assert item.ok  # нечего проверять


class TestCheckJson:
    def test_empty_data_returns_check_result(self):
        result = check_json({})
        assert isinstance(result, CheckResult)
        assert result.is_valid  # пустые данные — нет ошибок

    def test_valid_data_ok(self):
        data = {
            "general": {
                "ds_number": "3",
                "sum_text": 100.0,
                "probeg_sravnenie": 500.0,
            },
            "appendices": {
                "1": {"appendix_num": "1", "num_of_types": 0, "ds_num": "3"},
            },
        }
        result = check_json(data)
        assert isinstance(result, CheckResult)

    def test_sum_inconsistency_becomes_error(self):
        data = {
            "general": {"sum_text": 100.0, "sum_etapy": 200.0},
        }
        result = check_json(data)
        assert not result.is_valid
        assert any("sum_" in e for e in result.errors)
