import csv
from pathlib import Path
from typing import Any, cast

import pytest

from rfexcel.exception.library_exceptions import (
    FileDoesNotExistException, HeadersNotDeterminedException,
    StreamingViolationException, WorkbookNotOpenException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import *

XLS_FIRST_ROW = {
    "Index": 1, "First Name": "Dulce", "Last Name": "Abril",
    "Gender": "Female", "Country": "United States", "Age": 32,
}
XLS_LAST_ROW = {
    "Index": 9, "First Name": "Vincenza", "Last Name": "Weiland",
    "Gender": "Female", "Country": "United States", "Age": 40,
}

EXPECTED_ROW_COUNT: dict[str, int] = {
    "xlsx_edit": 4, "xlsx_stream": 4,
    "csv_edit":  4, "csv_stream":  4,
    "xls_edit":  9, "xls_on_demand": 9,
}

EXPECTED_ROWS: dict[str, list[Any]] = {
    "xlsx_edit":   XLSX_ROWS,
    "xlsx_stream": XLSX_ROWS,
    "csv_edit":    CSV_ROWS,
    "csv_stream":  CSV_ROWS,
}

EXPECTED_FIRST_ROW: dict[str, Any] = {
    "xlsx_edit": XLSX_ROWS[0],
    "csv_edit":  CSV_ROWS[0],
    "xls_edit":  XLS_FIRST_ROW,
}

EXPECTED_LAST_ROW: dict[str, Any] = {
    "xlsx_edit": XLSX_ROWS[-1],
    "csv_edit":  CSV_ROWS[-1],
    "xls_edit":  XLS_LAST_ROW,
}

EXACT_SEARCH_CASES: dict[str, tuple[dict[str, str], str, str]] = {
    "xlsx_edit": ({"Product ID": "P-200"}, "Product ID", "P-200"),
    "csv_edit":  ({"Product ID": "P-202"}, "Location",   "Paris, France"),
    "xls_edit":  ({"First Name": "Dulce"}, "Last Name",  "Abril"),
}

PARTIAL_SEARCH_CASES: dict[str, tuple[dict[str, str], str, str]] = {
    "xlsx_edit": ({"Description": "Keyboard"}, "Product ID",  "P-201"),
    "csv_edit":  ({"Description": "Keyboard"}, "Description", "Keyboard, Mechanical, RGB"),
    "xls_edit":  ({"Country":     "United"},   "First Name",  "Dulce"),
}

ONE_ROW_CASES: dict[str, tuple[dict[str, str] | None, str, str]] = {
    "xlsx_edit": (None,                    "Product ID", "P-200"),
    "csv_edit":  ({"Product ID": "P-203"}, "Description", "USB Cable, 3ft"),
    "xls_edit":  ({"First Name": "Dulce"}, "Last Name",   "Abril"),
}

STRING_CRITERIA_CASES: dict[str, tuple[str, dict[str, str]]] = {
    "xlsx_edit": ("Product ID=P-200",   {"Product ID": "P-200"}),
    "csv_edit":  ("Product ID=P-202",   {"Product ID": "P-202"}),
    "xls_edit":  ("First Name=Dulce",   {"First Name": "Dulce"}),
}

NO_MATCH_CRITERIA: dict[str, dict[str, str]] = {
    "xlsx_edit": {"Product ID": "NONEXISTENT"},
    "csv_edit":  {"Product ID": "NONEXISTENT"},
    "xls_edit":  {"First Name": "NONEXISTENT"},
}

NONEXISTENT_HEADER_CRITERIA: dict[str, dict[str, str]] = {
    "xlsx_edit": {"NonExistentColumn": "value"},
    "csv_edit":  {"NonExistentColumn": "value"},
    "xls_edit":  {"NonExistentColumn": "value"},
}

NUMERIC_INT_CRITERIA: dict[str, tuple[dict[str, str], str, str]] = {
    "xlsx_edit": ({"Price": "150"}, "Product ID", "P-202"),
    "csv_edit":  ({"Price": "150"}, "Product ID", "P-202"),
}

NUMERIC_FLOAT_CRITERIA: dict[str, tuple[dict[str, str], str, str]] = {
    "xlsx_edit": ({"Price": "25.5"}, "Product ID", "P-200"),
    "csv_edit":  ({"Price": "25.5"}, "Product ID", "P-200"),
}

COLUMN_HEADERS: dict[str, list[str]] = {
    "xlsx_edit": XLSX_HEADERS,
    "csv_edit":  XLSX_HEADERS,
    "xls_edit":  XLS_HEADERS_SHEET_1,
}



# ---------------------------------------------------------------------------
# Row count — all backends
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", list(BACKENDS))
def test_correct_row_count(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    assert len(lib.get_rows()) == EXPECTED_ROW_COUNT[backend_name]


# ---------------------------------------------------------------------------
# Full row content — xlsx and csv (edit + stream)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", list(EXPECTED_ROWS))
def test_all_rows_match_expected(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    assert lib.get_rows() == EXPECTED_ROWS[backend_name]


# ---------------------------------------------------------------------------
# First and last row — all edit backends
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", list(EXPECTED_FIRST_ROW))
def test_first_row_content(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    assert lib.get_rows()[0] == EXPECTED_FIRST_ROW[backend_name]


@pytest.mark.parametrize("backend_name", list(EXPECTED_LAST_ROW))
def test_last_row_content(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    assert lib.get_rows()[-1] == EXPECTED_LAST_ROW[backend_name]


# ---------------------------------------------------------------------------
# Stream parity — all three formats
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", list(FORMAT_FILE))
def test_stream_produces_same_as_edit_mode(lib: RFExcelLibrary, backend_name: str) -> None:
    path = FORMAT_FILE[backend_name]
    lib.load_workbook(path, read_only=True)
    stream_rows = lib.get_rows()
    lib.close()
    lib.load_workbook(path)
    assert lib.get_rows() == stream_rows


# ---------------------------------------------------------------------------
# Streaming violation — forward-only backends
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", ["xlsx_stream", "csv_stream"])
def test_calling_get_rows_twice_raises_streaming_violation(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    lib.get_rows()
    with pytest.raises(StreamingViolationException):
        lib.get_rows()


# ---------------------------------------------------------------------------
# Column header keys — all edit backends
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", list(COLUMN_HEADERS))
def test_all_rows_have_correct_header_keys(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    for row in cast(list[dict[str, Any]], lib.get_rows()):
        assert list(row.keys()) == COLUMN_HEADERS[backend_name]


# ---------------------------------------------------------------------------
# Header row out of range — xlsx and csv
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", ["xlsx_edit", "csv_edit"])
def test_header_row_out_of_range_raises(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    with pytest.raises(HeadersNotDeterminedException):
        lib.get_rows(header_row=9999)


# ---------------------------------------------------------------------------
# Custom header_row — xlsx
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", ["xlsx_edit"])
def test_default_header_row_equals_explicit_header_row_1(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    assert lib.get_rows() == lib.get_rows(header_row=1)


@pytest.mark.parametrize("backend_name", ["xlsx_edit"])
def test_header_row_2_shifts_data_by_one(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    rows = lib.get_rows(header_row=2)
    assert len(rows) == 3
    assert "P-200" in rows[0]


@pytest.mark.parametrize("backend_name", ["xlsx_edit"])
def test_header_row_beyond_data_returns_empty_list(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    assert lib.get_rows(header_row=5) == []


# ---------------------------------------------------------------------------
# Comma-containing values preserved — xlsx and csv
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", ["xlsx_edit", "csv_edit"])
def test_cell_containing_comma_is_not_split(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    rows = lib.get_rows()
    assert rows[0]["Location"] == "Warehouse A, Shelf 2"
    assert rows[2]["Location"] == "Paris, France"


@pytest.mark.parametrize("backend_name", ["csv_edit"])
def test_csv_quoted_field_with_comma_is_single_value(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    rows = lib.get_rows()
    assert rows[1]["Description"] == "Keyboard, Mechanical, RGB"


# ---------------------------------------------------------------------------
# XLS-specific: trailing empty columns excluded and numeric types cast to int
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", ["xls_edit"])
def test_xls_trailing_empty_columns_excluded_from_result(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    rows = lib.get_rows()
    assert "" not in rows[0]
    assert list(cast(dict[str, Any], rows[0]).keys()) == XLS_HEADERS_SHEET_1


@pytest.mark.parametrize("backend_name", ["xls_edit"])
def test_xls_whole_number_floats_are_cast_to_int(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    rows = lib.get_rows()
    assert isinstance(rows[0]["Index"], int)
    assert isinstance(rows[0]["Age"], int)
    assert rows[0]["Index"] == 1
    assert rows[0]["Age"] == 32


# ---------------------------------------------------------------------------
# Exact match search criteria — all edit backends
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", list(EXACT_SEARCH_CASES))
def test_exact_match_returns_one_matching_row(lib: RFExcelLibrary, backend_name: str) -> None:
    criteria, check_key, expected_val = EXACT_SEARCH_CASES[backend_name]
    open_backend(lib, backend_name)
    rows = lib.get_rows(search_criteria=criteria)
    assert len(rows) == 1
    assert rows[0][check_key] == expected_val


@pytest.mark.parametrize("backend_name", list(NO_MATCH_CRITERIA))
def test_exact_match_criteria_not_found_returns_empty(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    assert lib.get_rows(search_criteria=NO_MATCH_CRITERIA[backend_name]) == []


@pytest.mark.parametrize("backend_name", ["xlsx_edit", "csv_edit"])
def test_exact_match_full_value_required(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    assert lib.get_rows(search_criteria={"Description": "Keyboard"}) == []


# ---------------------------------------------------------------------------
# String-format criteria — all edit backends
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", list(STRING_CRITERIA_CASES))
def test_string_criteria_returns_same_as_dict(lib: RFExcelLibrary, backend_name: str) -> None:
    str_criteria, dict_criteria = STRING_CRITERIA_CASES[backend_name]
    open_backend(lib, backend_name)
    dict_rows = lib.get_rows(search_criteria=dict_criteria)
    str_rows  = lib.get_rows(search_criteria=str_criteria)
    assert dict_rows == str_rows


@pytest.mark.parametrize("backend_name", ["xlsx_edit", "csv_edit"])
def test_string_criteria_multiple_segments(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    rows = lib.get_rows(search_criteria="Product ID=P-200;Price=25.5")
    assert len(rows) == 1
    assert rows[0]["Product ID"] == "P-200"


@pytest.mark.parametrize("backend_name", ["xlsx_edit", "csv_edit", "xls_edit"])
def test_no_criteria_returns_all_rows(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    assert lib.get_rows() == lib.get_rows(search_criteria=None)


# ---------------------------------------------------------------------------
# AND logic — xlsx and csv
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", ["xlsx_edit", "csv_edit"])
def test_and_logic_two_criteria_narrows_result_to_one_row(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    rows = lib.get_rows(search_criteria={"Product ID": "P-202", "Price": "150"})
    assert len(rows) == 1
    assert rows[0]["Product ID"] == "P-202"


@pytest.mark.parametrize("backend_name", ["xlsx_edit", "csv_edit"])
def test_and_logic_conflicting_criteria_returns_empty(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    rows = lib.get_rows(search_criteria={"Product ID": "P-200", "Price": "150.00"})
    assert rows == []


# ---------------------------------------------------------------------------
# Numeric string criteria matches native types — xlsx and csv
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", list(NUMERIC_INT_CRITERIA))
def test_integer_string_criteria_matches_native_type(lib: RFExcelLibrary, backend_name: str) -> None:
    criteria, check_key, expected = NUMERIC_INT_CRITERIA[backend_name]
    open_backend(lib, backend_name)
    rows = lib.get_rows(search_criteria=criteria)
    assert len(rows) == 1
    assert rows[0][check_key] == expected


@pytest.mark.parametrize("backend_name", list(NUMERIC_FLOAT_CRITERIA))
def test_float_string_criteria_matches_native_type(lib: RFExcelLibrary, backend_name: str) -> None:
    criteria, check_key, expected = NUMERIC_FLOAT_CRITERIA[backend_name]
    open_backend(lib, backend_name)
    rows = lib.get_rows(search_criteria=criteria)
    assert len(rows) == 1
    assert rows[0][check_key] == expected


# ---------------------------------------------------------------------------
# Partial match — xlsx, csv, and xls
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", list(PARTIAL_SEARCH_CASES))
def test_partial_match_returns_matching_row(lib: RFExcelLibrary, backend_name: str) -> None:
    criteria, check_key, expected_val = PARTIAL_SEARCH_CASES[backend_name]
    open_backend(lib, backend_name)
    rows = lib.get_rows(search_criteria=criteria, partial_match=True)
    assert len(rows) >= 1
    assert rows[0][check_key] == expected_val


@pytest.mark.parametrize("backend_name", ["xlsx_edit", "csv_edit"])
def test_partial_match_false_does_not_match_substring(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    assert lib.get_rows(search_criteria={"Description": "Keyboard"}, partial_match=False) == []


@pytest.mark.parametrize("backend_name", ["xlsx_edit", "csv_edit"])
def test_partial_match_and_logic_both_criteria_must_match(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    rows = lib.get_rows(search_criteria={"Description": "Mouse", "Location": "Warehouse"}, partial_match=True)
    assert len(rows) == 1
    assert rows[0]["Product ID"] == "P-200"


@pytest.mark.parametrize("backend_name", ["xlsx_edit"])
def test_partial_match_substring_location(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    rows = lib.get_rows(search_criteria={"Location": "France"}, partial_match=True)
    assert len(rows) == 1
    assert rows[0]["Product ID"] == "P-202"


@pytest.mark.parametrize("backend_name", ["xls_edit"])
def test_partial_match_returns_multiple_rows(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    rows = lib.get_rows(search_criteria={"Country": "United"}, partial_match=True)
    assert len(rows) == 6
    assert all("United" in r["Country"] for r in rows)


# ---------------------------------------------------------------------------
# Nonexistent criteria key returns empty — all edit backends
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", list(NONEXISTENT_HEADER_CRITERIA))
def test_criteria_key_not_in_headers_returns_empty(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    assert lib.get_rows(search_criteria=NONEXISTENT_HEADER_CRITERIA[backend_name]) == []


# ---------------------------------------------------------------------------
# Search criteria in xls stream (on-demand) mode
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", ["xls_on_demand"])
def test_criteria_on_xls_stream_mode(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    rows = lib.get_rows(search_criteria={"First Name": "Dulce"})
    assert len(rows) == 1
    assert rows[0]["First Name"] == "Dulce"


# ---------------------------------------------------------------------------
# one_row=True — all edit backends
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("backend_name", list(ONE_ROW_CASES))
def test_one_row_returns_dict_with_correct_data(lib: RFExcelLibrary, backend_name: str) -> None:
    criteria, check_key, expected_val = ONE_ROW_CASES[backend_name]
    open_backend(lib, backend_name)
    result = lib.get_rows(search_criteria=criteria, one_row=True)
    assert isinstance(result, dict)
    assert result[check_key] == expected_val


@pytest.mark.parametrize("backend_name", ["xlsx_edit", "csv_edit", "xls_edit"])
def test_one_row_no_match_returns_empty_dict(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    assert lib.get_rows(search_criteria=NO_MATCH_CRITERIA[backend_name], one_row=True) == {}


@pytest.mark.parametrize("backend_name", ["xlsx_edit"])
def test_one_row_with_partial_match(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    result = lib.get_rows(search_criteria={"Description": "Keyboard"}, partial_match=True, one_row=True)
    assert isinstance(result, dict)
    assert result["Product ID"] == "P-201"


@pytest.mark.parametrize("backend_name", ["xlsx_edit", "csv_edit", "xls_edit"])
def test_one_row_false_returns_list(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    result = lib.get_rows(one_row=False)
    assert isinstance(result, list)
    assert len(result) == EXPECTED_ROW_COUNT[backend_name]


@pytest.mark.parametrize("backend_name", ["xlsx_edit"])
def test_one_row_early_exit_does_not_exhaust_all_rows(lib: RFExcelLibrary, backend_name: str) -> None:
    open_backend(lib, backend_name)
    first = lib.get_rows(one_row=True)
    assert first == XLSX_ROWS[0]
    all_rows = lib.get_rows()
    assert len(all_rows) == 4


# ---------------------------------------------------------------------------
# Negative / error cases
# ---------------------------------------------------------------------------

def test_raises_when_no_workbook_loaded(lib: RFExcelLibrary) -> None:
    with pytest.raises(WorkbookNotOpenException):
        lib.get_rows()


def test_raises_after_close(lib: RFExcelLibrary) -> None:
    lib.load_workbook(XLSX_FILE)
    lib.close()
    with pytest.raises(WorkbookNotOpenException):
        lib.get_rows()


def test_one_row_no_workbook_raises(lib: RFExcelLibrary) -> None:
    with pytest.raises(WorkbookNotOpenException):
        lib.get_rows(one_row=True)


def test_load_nonexistent_file_raises(lib: RFExcelLibrary) -> None:
    with pytest.raises(FileDoesNotExistException):
        lib.load_workbook("/nonexistent/path/missing.xlsx")


def test_get_rows_on_empty_created_xlsx_returns_empty_list(lib: RFExcelLibrary, tmp_path: Path) -> None:
    lib.create_workbook(str(tmp_path / "empty.xlsx"))
    assert lib.get_rows() == []


def test_get_rows_on_empty_created_csv_raises(lib: RFExcelLibrary, tmp_path: Path) -> None:
    lib.create_workbook(str(tmp_path / "empty.csv"))
    with pytest.raises(HeadersNotDeterminedException):
        lib.get_rows()


def test_header_row_1_on_csv_with_only_headers_returns_empty_list(lib: RFExcelLibrary, tmp_path: Path) -> None:
    path = tmp_path / "headers_only.csv"
    with open(path, "w", newline="") as f:
        csv.writer(f).writerow(["A", "B", "C"])
    lib.load_workbook(str(path))
    assert lib.get_rows() == []
