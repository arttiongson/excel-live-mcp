"""
Unit tests for excel_mcp.client. No live Excel required, uses mocks.

Run:  pytest tests/ -v
"""
from __future__ import annotations

import unittest.mock as mock

import pandas as pd
import pytest

from excel_mcp import client as ec


# ---------- _excel_col_letter_to_num ----------
class TestExcelColLetter:
    def test_single_letter(self):
        assert ec._excel_col_letter_to_num("A") == 1
        assert ec._excel_col_letter_to_num("B") == 2
        assert ec._excel_col_letter_to_num("Z") == 26

    def test_double_letter(self):
        assert ec._excel_col_letter_to_num("AA") == 27
        assert ec._excel_col_letter_to_num("AB") == 28
        assert ec._excel_col_letter_to_num("AZ") == 52
        assert ec._excel_col_letter_to_num("BA") == 53
        assert ec._excel_col_letter_to_num("ZZ") == 702

    def test_triple_letter(self):
        assert ec._excel_col_letter_to_num("AAA") == 703
        assert ec._excel_col_letter_to_num("XFD") == 16384

    def test_lowercase(self):
        assert ec._excel_col_letter_to_num("aa") == 27


# ---------- _resolve_col ----------
class TestResolveCol:
    def test_exact_name(self):
        df = pd.DataFrame({"Item": [1, 2], "Qty": [3, 4]})
        assert ec._resolve_col(df, "Item") == "Item"

    def test_case_insensitive(self):
        df = pd.DataFrame({"Item": [1], "Qty": [2]})
        assert ec._resolve_col(df, "item") == "Item"
        assert ec._resolve_col(df, "QTY") == "Qty"

    def test_letter_at_origin(self):
        df = pd.DataFrame({"Item": [1], "Qty": [2], "Cost": [3]})
        assert ec._resolve_col(df, "A", used_col=1) == "Item"
        assert ec._resolve_col(df, "B", used_col=1) == "Qty"
        assert ec._resolve_col(df, "C", used_col=1) == "Cost"

    def test_letter_with_used_col_offset(self):
        """REGRESSION: used_range starts at column C, 'C' maps to df.columns[0]."""
        df = pd.DataFrame({"Item": [1], "Qty": [2], "Cost": [3]})
        assert ec._resolve_col(df, "C", used_col=3) == "Item"
        assert ec._resolve_col(df, "D", used_col=3) == "Qty"
        assert ec._resolve_col(df, "E", used_col=3) == "Cost"

    def test_letter_outside_used_range_raises(self):
        """REGRESSION: 'A' against used_col=3 (range starts at C) must error."""
        df = pd.DataFrame({"Item": [1], "Qty": [2]})
        with pytest.raises(ec.ExcelError, match="outside used range"):
            ec._resolve_col(df, "A", used_col=3)

    def test_letter_multi_char(self):
        df = pd.DataFrame({c: [1] for c in ["First", "Second", "Third", "Fourth", "Fifth"]})
        with pytest.raises(ec.ExcelError, match="outside used range"):
            ec._resolve_col(df, "AA", used_col=1)
        assert ec._resolve_col(df, "AA", used_col=23) == "Fifth"

    def test_mixed_case_names_not_treated_as_letters(self):
        """REGRESSION: 'Nonexistent' must not be parsed as a column letter."""
        df = pd.DataFrame({"Item": [1]})
        with pytest.raises(ec.ExcelError, match="not found"):
            ec._resolve_col(df, "Nonexistent")

    def test_lowercase_letter_falls_through(self):
        df = pd.DataFrame({"Item": [1]})
        with pytest.raises(ec.ExcelError, match="not found"):
            ec._resolve_col(df, "aa")

    def test_unknown_column_raises(self):
        df = pd.DataFrame({"Item": [1]})
        with pytest.raises(ec.ExcelError, match="not found"):
            ec._resolve_col(df, "Whatever123")


# ---------- _build_equals_mask ----------
class TestEqualsMask:
    def test_int_against_float_column(self):
        """REGRESSION: Excel reads ints as floats. equals=42 must match 42.0."""
        s = pd.Series([42.0, 42, 43.0, "42", None])
        mask = ec._build_equals_mask(s, 42)
        assert mask.tolist() == [True, True, False, True, False]

    def test_float_against_int_column(self):
        s = pd.Series([100, 100.0, 200.5])
        mask = ec._build_equals_mask(s, 100.0)
        assert mask.tolist() == [True, True, False]

    def test_zero(self):
        s = pd.Series([0, 0.0, "0", None, 1])
        mask = ec._build_equals_mask(s, 0)
        assert mask.tolist() == [True, True, True, False, False]

    def test_negative_number(self):
        s = pd.Series([-5, -5.0, 5, "-5"])
        mask = ec._build_equals_mask(s, -5)
        assert mask.tolist() == [True, True, False, True]

    def test_string_exact(self):
        s = pd.Series(["apple", "Apple", "banana"])
        mask = ec._build_equals_mask(s, "apple")
        assert mask.tolist() == [True, False, False]

    def test_bool_exact(self):
        s = pd.Series([True, False, True])
        mask = ec._build_equals_mask(s, True)
        assert mask.tolist() == [True, False, True]

    def test_null_in_column(self):
        s = pd.Series([1.0, None, 2.0, float("nan")])
        mask = ec._build_equals_mask(s, 1)
        assert mask.tolist() == [True, False, False, False]


# ---------- _build_contains_mask ----------
class TestContainsMask:
    def test_case_insensitive(self):
        s = pd.Series(["Apple Pie", "banana", "PINEAPPLE", None])
        mask = ec._build_contains_mask(s, "apple")
        assert mask.tolist() == [True, False, True, False]

    def test_regex_chars_escaped(self):
        """REGRESSION: '.' must match literal '.', not 'any char'."""
        s = pd.Series(["foo.bar", "fooXbar", "foo"])
        mask = ec._build_contains_mask(s, ".")
        assert mask.tolist() == [True, False, False]

    def test_empty_substring(self):
        s = pd.Series(["abc", "", None])
        mask = ec._build_contains_mask(s, "")
        assert mask.tolist()[0] is True


# ---------- _clean_for_json ----------
class TestCleanForJson:
    def test_nan_to_none(self):
        assert ec._clean_for_json(float("nan")) is None

    def test_inf_to_none(self):
        assert ec._clean_for_json(float("inf")) is None
        assert ec._clean_for_json(float("-inf")) is None

    def test_normal_float_preserved(self):
        assert ec._clean_for_json(42.0) == 42.0
        assert ec._clean_for_json(0.0) == 0.0
        assert ec._clean_for_json(-1.5) == -1.5

    def test_dict_recursive(self):
        obj = {"a": 1, "b": float("nan"), "c": "ok"}
        assert ec._clean_for_json(obj) == {"a": 1, "b": None, "c": "ok"}

    def test_nested_groupby_shape(self):
        obj = {"groups": [{"key": "x", "sum": float("nan")}, {"key": "y", "sum": 3.0}]}
        out = ec._clean_for_json(obj)
        assert out == {"groups": [{"key": "x", "sum": None}, {"key": "y", "sum": 3.0}]}

    def test_list(self):
        assert ec._clean_for_json([1, float("nan"), 3]) == [1, None, 3]

    def test_pd_na(self):
        assert ec._clean_for_json(pd.NA) is None

    def test_strings_preserved(self):
        assert ec._clean_for_json("hello") == "hello"
        assert ec._clean_for_json("nan") == "nan"


# ---------- _normalize_2d ----------
class TestNormalize2d:
    def test_1x1_scalar(self):
        assert ec._normalize_2d(42, (1, 1)) == [[42]]

    def test_1xN_single_row(self):
        assert ec._normalize_2d([1, 2, 3], (1, 3)) == [[1, 2, 3]]

    def test_Nx1_single_column(self):
        assert ec._normalize_2d([1, 2, 3], (3, 1)) == [[1], [2], [3]]

    def test_NxM_already_2d(self):
        assert ec._normalize_2d([[1, 2], [3, 4]], (2, 2)) == [[1, 2], [3, 4]]

    def test_1x1_none(self):
        assert ec._normalize_2d(None, (1, 1)) == [[None]]

    def test_NxM_none(self):
        assert ec._normalize_2d(None, (2, 3)) == [[None, None, None], [None, None, None]]


# ---------- Tool-level: count_where ----------
class TestCountWhere:
    def test_no_predicate_raises(self):
        df = pd.DataFrame({"x": [1]})
        with mock.patch.object(ec, "_get_sheet"), mock.patch.object(
            ec, "_read_sheet_cached", return_value=(df, 1, 1)
        ):
            with pytest.raises(ec.ExcelError, match="pass one of"):
                ec.count_where("wb", "sht", "x")

    def test_multiple_predicates_raises(self):
        df = pd.DataFrame({"x": [1]})
        with mock.patch.object(ec, "_get_sheet"), mock.patch.object(
            ec, "_read_sheet_cached", return_value=(df, 1, 1)
        ):
            with pytest.raises(ec.ExcelError, match="only ONE"):
                ec.count_where("wb", "sht", "x", equals=1, contains="foo")

    def test_numeric_equals_matches_excel_floats(self):
        """REGRESSION: equals=100 on a column where xlwings returned 100.0 must match."""
        df = pd.DataFrame({"Qty": [100.0, 100, 200, "100", None]})
        with mock.patch.object(ec, "_get_sheet"), mock.patch.object(
            ec, "_read_sheet_cached", return_value=(df, 1, 1)
        ):
            result = ec.count_where("wb", "sht", "Qty", equals=100)
        assert result["matches"] == 3
        assert result["total_rows"] == 5

    def test_not_null(self):
        df = pd.DataFrame({"x": [1, None, 3, None, 5]})
        with mock.patch.object(ec, "_get_sheet"), mock.patch.object(
            ec, "_read_sheet_cached", return_value=(df, 1, 1)
        ):
            result = ec.count_where("wb", "sht", "x", not_null=True)
        assert result["matches"] == 3


# ---------- Tool-level: groupby_sum ----------
class TestGroupbySum:
    def test_nan_key_serializes_as_none(self):
        """REGRESSION: null group keys come back as None, not 'nan' string."""
        df = pd.DataFrame({"Cat": ["A", "B", None, "A"], "Val": [10, 20, 30, 40]})
        with mock.patch.object(ec, "_get_sheet"), mock.patch.object(
            ec, "_read_sheet_cached", return_value=(df, 1, 1)
        ):
            result = ec.groupby_sum("wb", "sht", "Cat", "Val")
        keys = [g["key"] for g in result["groups"]]
        assert None in keys
        assert "nan" not in keys

    def test_grand_total(self):
        df = pd.DataFrame({"Cat": ["A", "B", "A"], "Val": [10, 20, 30]})
        with mock.patch.object(ec, "_get_sheet"), mock.patch.object(
            ec, "_read_sheet_cached", return_value=(df, 1, 1)
        ):
            result = ec.groupby_sum("wb", "sht", "Cat", "Val")
        assert result["grand_total"] == 60.0

    def test_string_numbers_coerced(self):
        df = pd.DataFrame({"Cat": ["A", "A"], "Val": ["10", "20"]})
        with mock.patch.object(ec, "_get_sheet"), mock.patch.object(
            ec, "_read_sheet_cached", return_value=(df, 1, 1)
        ):
            result = ec.groupby_sum("wb", "sht", "Cat", "Val")
        assert result["grand_total"] == 30.0


# ---------- Tool-level: find_rows ----------
class TestFindRows:
    def test_no_predicate_raises(self):
        df = pd.DataFrame({"x": [1]})
        with mock.patch.object(ec, "_get_sheet"), mock.patch.object(
            ec, "_read_sheet_cached", return_value=(df, 1, 1)
        ):
            with pytest.raises(ec.ExcelError, match="pass equals OR contains"):
                ec.find_rows("wb", "sht", "x")

    def test_multiple_predicates_raises(self):
        df = pd.DataFrame({"x": [1]})
        with mock.patch.object(ec, "_get_sheet"), mock.patch.object(
            ec, "_read_sheet_cached", return_value=(df, 1, 1)
        ):
            with pytest.raises(ec.ExcelError, match="only ONE"):
                ec.find_rows("wb", "sht", "x", equals=1, contains="foo")

    def test_sheet_row_no_offset(self):
        df = pd.DataFrame({"Item": ["A", "B", "C"]})
        with mock.patch.object(ec, "_get_sheet"), mock.patch.object(
            ec, "_read_sheet_cached", return_value=(df, 1, 1)
        ):
            result = ec.find_rows("wb", "sht", "Item", equals="A")
        assert result["sheet_rows"] == [2]

    def test_sheet_row_with_offset(self):
        """REGRESSION: used_range starts at row 5, df.index 1 -> sheet row 7."""
        df = pd.DataFrame({"Item": ["A", "B", "C"]})
        with mock.patch.object(ec, "_get_sheet"), mock.patch.object(
            ec, "_read_sheet_cached", return_value=(df, 5, 1)
        ):
            result = ec.find_rows("wb", "sht", "Item", equals="B")
        assert result["sheet_rows"] == [7]

    def test_sheet_row_no_header(self):
        df = pd.DataFrame({0: ["A", "B", "C"]})
        with mock.patch.object(ec, "_get_sheet"), mock.patch.object(
            ec, "_read_sheet_cached", return_value=(df, 1, 1)
        ):
            result = ec.find_rows("wb", "sht", "A", equals="A", has_header=False)
        assert result["sheet_rows"] == [1]

    def test_numeric_match_does_not_silently_miss(self):
        """REGRESSION: equals=100 on column read as 100.0 must return the match."""
        df = pd.DataFrame({"Qty": [50.0, 100.0, 150.0]})
        with mock.patch.object(ec, "_get_sheet"), mock.patch.object(
            ec, "_read_sheet_cached", return_value=(df, 1, 1)
        ):
            result = ec.find_rows("wb", "sht", "Qty", equals=100)
        assert result["match_count"] == 1
        assert result["sheet_rows"] == [3]


# ---------- Tool-level: column_stats ----------
class TestColumnStats:
    def test_numeric_column(self):
        df = pd.DataFrame({"Val": [10, 20, 30, 40, 50]})
        with mock.patch.object(ec, "_get_sheet"), mock.patch.object(
            ec, "_read_sheet_cached", return_value=(df, 1, 1)
        ):
            result = ec.column_stats("wb", "sht", "Val")
        assert result["is_numeric"] is True
        assert result["min"] == 10.0
        assert result["max"] == 50.0
        assert result["mean"] == 30.0
        assert result["sum"] == 150.0

    def test_text_column(self):
        df = pd.DataFrame({"Cat": ["A", "B", "A", "C", "A"]})
        with mock.patch.object(ec, "_get_sheet"), mock.patch.object(
            ec, "_read_sheet_cached", return_value=(df, 1, 1)
        ):
            result = ec.column_stats("wb", "sht", "Cat")
        assert result["is_numeric"] is False
        assert result["unique"] == 3
        assert result["top_5"][0] == {"value": "A", "count": 3}

    def test_all_nulls(self):
        df = pd.DataFrame({"Val": [None, None, None]})
        with mock.patch.object(ec, "_get_sheet"), mock.patch.object(
            ec, "_read_sheet_cached", return_value=(df, 1, 1)
        ):
            result = ec.column_stats("wb", "sht", "Val")
        assert result["non_null"] == 0
        assert result["null"] == 3


# ---------- Tool-level: unique_values ----------
class TestUniqueValues:
    def test_basic(self):
        df = pd.DataFrame({"x": ["a", "b", "a", "c", None]})
        with mock.patch.object(ec, "_get_sheet"), mock.patch.object(
            ec, "_read_sheet_cached", return_value=(df, 1, 1)
        ):
            result = ec.unique_values("wb", "sht", "x")
        assert set(result["values"]) == {"a", "b", "c"}
        assert result["count"] == 3
        assert result["truncated"] is False
