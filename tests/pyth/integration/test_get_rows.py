import csv
from pathlib import Path
from typing import Any, cast

import pytest

from rfexcel.exception.library_exceptions import (
    FileDoesNotExistException, HeadersNotDeterminedException,
    StreamingViolationException, WorkbookNotOpenException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

XLSX_HEADERS = ["Product ID", "Description", "Price", "Location"]

XLSX_ROWS = [
    {"Product ID": "P-200", "Description": "Wireless Mouse",            "Price": 25.5,  "Location": "Warehouse A, Shelf 2"},
    {"Product ID": "P-201", "Description": "Keyboard, Mechanical",      "Price": 89.99, "Location": "Store Front"},
    {"Product ID": "P-202", "Description": "Monitor 24-inch",           "Price": 150,   "Location": "Paris, France"},
    {"Product ID": "P-203", "Description": "USB Cable",                 "Price": 5.99,  "Location": "OnlineP"},
]

CSV_ROWS = [
    {"Product ID": "P-200", "Description": "Wireless Mouse",            "Price": 25.5,   "Location": "Warehouse A, Shelf 2"},
    {"Product ID": "P-201", "Description": "Keyboard, Mechanical, RGB", "Price": 89.99,  "Location": "Store Front"},
    {"Product ID": "P-202", "Description": "Monitor 24-inch",           "Price": 150,    "Location": "Paris, France"},
    {"Product ID": "P-203", "Description": "USB Cable, 3ft",            "Price": 5.99,   "Location": "Online"},
]

XLS_FIRST_ROW = {
    "Index": 1.0, "First Name": "Dulce", "Last Name": "Abril",
    "Gender": "Female", "Country": "United States", "Age": 32.0,
}
XLS_LAST_ROW = {
    "Index": 9.0, "First Name": "Vincenza", "Last Name": "Weiland",
    "Gender": "Female", "Country": "United States", "Age": 40.0,
}


# ---------------------------------------------------------------------------
# Row count – all formats and modes
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    ("path", "read_only", "expected_count"),
    [
        (XLSX_FILE, False, 4),
        (XLSX_FILE, True,  4),
        (XLS_FILE,  False, 9),
        (XLS_FILE,  True,  9),
        (CSV_FILE,  False, 4),
        (CSV_FILE,  True,  4),
    ],
    ids=["xlsx_edit", "xlsx_stream", "xls_edit", "xls_on_demand", "csv_edit", "csv_stream"],
)
def test_correct_row_count(
    lib: RFExcelLibrary, path: str, read_only: bool, expected_count: int
):
    lib.load_workbook(path, read_only=read_only)
    assert len(lib.get_rows()) == expected_count


# ---------------------------------------------------------------------------
# Row content – formats with identical expected data across modes
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    ("path", "read_only", "expected"),
    [
        (XLSX_FILE, False, XLSX_ROWS),
        (XLSX_FILE, True,  XLSX_ROWS),
        (CSV_FILE,  False, CSV_ROWS),
        (CSV_FILE,  True,  CSV_ROWS),
    ],
    ids=["xlsx_edit", "xlsx_stream", "csv_edit", "csv_stream"],
)
def test_all_rows_match_expected(
    lib: RFExcelLibrary, path: str, read_only: bool, expected: list[Any]
):
    lib.load_workbook(path, read_only=read_only)
    assert lib.get_rows() == expected


# ---------------------------------------------------------------------------
# Stream mode parity – all three backends
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("path", [XLSX_FILE, XLS_FILE, CSV_FILE], ids=["xlsx", "xls", "csv"])
def test_stream_produces_same_as_edit_mode(lib: RFExcelLibrary, path: str):
    lib.load_workbook(path, read_only=True)
    stream_rows = lib.get_rows()
    lib.close()
    lib.load_workbook(path)
    assert lib.get_rows() == stream_rows


@pytest.mark.parametrize("path", [XLSX_FILE, CSV_FILE], ids=["xlsx", "csv"])
def test_calling_get_rows_twice_raises_streaming_violation(lib: RFExcelLibrary, path: str):
    lib.load_workbook(path, read_only=True)
    lib.get_rows()
    with pytest.raises(StreamingViolationException):
        lib.get_rows()


# ---------------------------------------------------------------------------
# Shared header behaviour – xlsx and csv
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("path", [XLSX_FILE, CSV_FILE], ids=["xlsx", "csv"])
def test_all_rows_have_correct_header_keys(lib: RFExcelLibrary, path: str):
    lib.load_workbook(path)
    for row in cast(list[dict[str, Any]], lib.get_rows()):
        assert list(row.keys()) == XLSX_HEADERS


@pytest.mark.parametrize("path", [XLSX_FILE, CSV_FILE], ids=["xlsx", "csv"])
def test_header_row_out_of_range_raises(lib: RFExcelLibrary, path: str):
    lib.load_workbook(path)
    with pytest.raises(HeadersNotDeterminedException):
        lib.get_rows(header_row=9999)


# ---------------------------------------------------------------------------
# Search criteria – cross-backend exact and partial match
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    ("path", "criteria", "check_key", "expected_val"),
    [
        (XLSX_FILE, {"Product ID": "P-200"}, "Product ID", "P-200"),
        (CSV_FILE,  {"Product ID": "P-202"}, "Location",   "Paris, France"),
        (XLS_FILE,  {"First Name": "Dulce"}, "Last Name",  "Abril"),
    ],
    ids=["xlsx", "csv", "xls"],
)
def test_exact_match_returns_one_matching_row(
    lib: RFExcelLibrary, path: str, criteria: dict[str, str], check_key: str, expected_val: str
):
    lib.load_workbook(path)
    rows = lib.get_rows(search_criteria=criteria)
    assert len(rows) == 1
    assert rows[0][check_key] == expected_val


@pytest.mark.parametrize(
    ("path", "criteria", "check_key", "expected_val"),
    [
        (XLSX_FILE, {"Description": "Keyboard"}, "Product ID",  "P-201"),
        (CSV_FILE,  {"Description": "Keyboard"}, "Description", "Keyboard, Mechanical, RGB"),
    ],
    ids=["xlsx", "csv"],
)
def test_partial_match_returns_one_row(
    lib: RFExcelLibrary, path: str, criteria: dict[str, str], check_key: str, expected_val: str
):
    lib.load_workbook(path)
    rows = lib.get_rows(search_criteria=criteria, partial_match=True)
    assert len(rows) == 1
    assert rows[0][check_key] == expected_val


# ---------------------------------------------------------------------------
# one_row=True – cross-backend
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    ("path", "criteria", "check_key", "expected_val"),
    [
        (XLSX_FILE, None,                    "Product ID",  "P-200"),
        (CSV_FILE,  {"Product ID": "P-203"}, "Description", "USB Cable, 3ft"),
        (XLS_FILE,  {"First Name": "Dulce"}, "Last Name",   "Abril"),
    ],
    ids=["xlsx", "csv", "xls"],
)
def test_one_row_returns_dict_with_correct_data(
    lib: RFExcelLibrary,
    path: str,
    criteria: dict[str, str] | None,
    check_key: str,
    expected_val: str,
):
    lib.load_workbook(path)
    result = lib.get_rows(search_criteria=criteria, one_row=True)
    assert isinstance(result, dict)
    assert result[check_key] == expected_val


# ---------------------------------------------------------------------------
# xlsx edit – format-specific behaviour
# ---------------------------------------------------------------------------

class TestGetRowsXlsxEdit:

    def test_cell_containing_comma_is_not_split(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows()
        assert rows[0]["Location"] == "Warehouse A, Shelf 2"
        assert rows[2]["Location"] == "Paris, France"

    def test_default_header_row_equals_explicit_header_row_1(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert lib.get_rows() == lib.get_rows(header_row=1)

    def test_header_row_2_shifts_data_by_one(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows(header_row=2)
        assert len(rows) == 3
        assert "P-200" in rows[0]

    def test_header_row_beyond_data_returns_empty_list(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert lib.get_rows(header_row=5) == []


# ---------------------------------------------------------------------------
# csv edit – format-specific behaviour
# ---------------------------------------------------------------------------

def test_csv_quoted_field_with_comma_is_single_value(lib: RFExcelLibrary):
    lib.load_workbook(CSV_FILE)
    rows = lib.get_rows()
    assert rows[1]["Description"] == "Keyboard, Mechanical, RGB"
    assert rows[0]["Location"] == "Warehouse A, Shelf 2"


# ---------------------------------------------------------------------------
# xls standard – format-specific behaviour
# ---------------------------------------------------------------------------

class TestGetRowsXlsStandard:

    def test_first_row_content(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        assert lib.get_rows()[0] == XLS_FIRST_ROW

    def test_last_row_content(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        assert lib.get_rows()[-1] == XLS_LAST_ROW

    def test_numeric_values_are_native_floats(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        rows = lib.get_rows()
        assert rows[0]["Index"] == 1.0
        assert rows[0]["Age"] == 32.0

    def test_trailing_empty_columns_excluded_from_result(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        rows = lib.get_rows()
        assert "" not in rows[0]
        assert list(cast(dict[str, Any], rows[0]).keys()) == ["Index", "First Name", "Last Name", "Gender", "Country", "Age"]

    def test_all_rows_contain_expected_name_columns(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        for row in lib.get_rows():
            assert "First Name" in row
            assert "Last Name" in row
            assert "Country" in row


# ---------------------------------------------------------------------------
# negative
# ---------------------------------------------------------------------------

class TestGetRowsNegative:

    def test_raises_when_no_workbook_loaded(self, lib: RFExcelLibrary):
        with pytest.raises(WorkbookNotOpenException):
            lib.get_rows()

    def test_raises_after_close(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.close()
        with pytest.raises(WorkbookNotOpenException):
            lib.get_rows()

    def test_load_nonexistent_file_raises(self, lib: RFExcelLibrary):
        with pytest.raises(FileDoesNotExistException):
            lib.load_workbook("/nonexistent/path/missing.xlsx")

    def test_get_rows_on_empty_created_xlsx_returns_empty_list(self, lib: RFExcelLibrary, tmp_path: Path):
        lib.create_workbook(str(tmp_path / "empty.xlsx"))
        assert lib.get_rows() == []

    def test_get_rows_on_empty_created_csv_raises(self, lib: RFExcelLibrary, tmp_path: Path):
        lib.create_workbook(str(tmp_path / "empty.csv"))
        with pytest.raises(HeadersNotDeterminedException):
            lib.get_rows()

    def test_header_row_1_on_single_row_file_returns_empty_list(self, lib: RFExcelLibrary, tmp_path: Path):
        path = tmp_path / "headers_only.csv"
        with open(path, "w", newline="") as f:
            csv.writer(f).writerow(["A", "B", "C"])
        lib.load_workbook(str(path))
        assert lib.get_rows() == []


# ---------------------------------------------------------------------------
# search criteria
# ---------------------------------------------------------------------------

class TestGetRowsSearchCriteria:

    def test_exact_match_dict_returns_correct_full_row(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows(search_criteria={"Product ID": "P-202"})
        assert rows == [XLSX_ROWS[2]]

    def test_exact_match_criteria_not_found_returns_empty(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert lib.get_rows(search_criteria={"Product ID": "NONEXISTENT"}) == []

    def test_exact_match_full_value_required(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert lib.get_rows(search_criteria={"Description": "Keyboard"}) == []

    def test_string_criteria_returns_same_as_dict(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        dict_rows = lib.get_rows(search_criteria={"Product ID": "P-200"})
        str_rows  = lib.get_rows(search_criteria="Product ID=P-200")
        assert dict_rows == str_rows

    def test_string_criteria_multiple_segments(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows(search_criteria="Product ID=P-200;Price=25.5")
        assert len(rows) == 1
        assert rows[0]["Product ID"] == "P-200"

    def test_string_criteria_no_criteria_returns_all(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert lib.get_rows() == lib.get_rows(search_criteria=None)

    def test_and_logic_two_criteria_narrows_result_to_one_row(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows(search_criteria={"Product ID": "P-202", "Price": "150"})
        assert len(rows) == 1
        assert rows[0]["Product ID"] == "P-202"

    def test_exact_match_accepts_non_string_value_in_dict_criteria(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows(search_criteria={"Price": "150"})
        assert len(rows) == 1
        assert rows[0]["Product ID"] == "P-202"

    def test_dict_search_criteria_float_matches_native_xlsx_type(self, lib: RFExcelLibrary):
        """Search with float 25.5 matches XLSX native float type."""
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows(search_criteria={"Price": "25.5"})
        assert len(rows) == 1
        assert rows[0]["Product ID"] == "P-200"

    def test_and_logic_conflicting_criteria_returns_empty(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows(search_criteria={"Product ID": "P-200", "Price": "150.00"})
        assert rows == []

    def test_partial_match_true_location_substring(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows(search_criteria={"Location": "France"}, partial_match=True)
        assert len(rows) == 1
        assert rows[0]["Product ID"] == "P-202"

    def test_partial_match_false_does_not_match_substring(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows(search_criteria={"Description": "Keyboard"}, partial_match=False)
        assert rows == []

    def test_partial_match_and_logic_both_criteria_must_match(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows(search_criteria={"Description": "Mouse", "Location": "Warehouse"}, partial_match=True)
        assert len(rows) == 1
        assert rows[0]["Product ID"] == "P-200"

    def test_partial_match_on_xls_multiple_rows(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        rows = lib.get_rows(search_criteria={"Country": "United"}, partial_match=True)
        assert len(rows) == 6
        assert all("United" in r["Country"] for r in rows)

    def test_criteria_key_not_in_headers_returns_empty(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert lib.get_rows(search_criteria={"NonExistentColumn": "value"}) == []

    def test_criteria_on_xls_stream_mode(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE, read_only=True)
        rows = lib.get_rows(search_criteria={"First Name": "Dulce"})
        assert len(rows) == 1
        assert rows[0]["First Name"] == "Dulce"


# ---------------------------------------------------------------------------
# one row
# ---------------------------------------------------------------------------

class TestGetRowsOneRow:

    def test_one_row_returns_first_row(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        result = lib.get_rows(one_row=True)
        assert result == XLSX_ROWS[0]

    def test_one_row_with_search_criteria_returns_matching_row(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        result = lib.get_rows(search_criteria={"Product ID": "P-202"}, one_row=True)
        assert result == XLSX_ROWS[2]

    def test_one_row_no_match_returns_empty_dict(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        result = lib.get_rows(search_criteria={"Product ID": "NOPE"}, one_row=True)
        assert result == {}

    def test_one_row_with_partial_match(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        result = lib.get_rows(search_criteria={"Description": "Keyboard"}, partial_match=True, one_row=True)
        assert isinstance(result, dict)
        assert result["Product ID"] == "P-201"

    def test_one_row_false_returns_list(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        result = lib.get_rows(one_row=False)
        assert isinstance(result, list)
        assert len(result) == 4

    def test_one_row_no_workbook_raises(self, lib: RFExcelLibrary):
        with pytest.raises(WorkbookNotOpenException):
            lib.get_rows(one_row=True)

    def test_one_row_early_exit_does_not_exhaust_all_rows(self, lib: RFExcelLibrary, tmp_path: Path):
        lib.load_workbook(XLSX_FILE)
        first = lib.get_rows(one_row=True)
        assert first == XLSX_ROWS[0]
        all_rows = lib.get_rows()
        assert len(all_rows) == 4
