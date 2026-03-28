from pathlib import Path
import csv

import pytest

from rfexcel.exception.library_exceptions import (
    FileDoesNotExistException, HeadersNotDeterminedException,
    StreamingViolationException, WorkbookNotOpenException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE


XLSX_HEADERS = ["Product ID", "Description", "Price", "Location"]

XLSX_ROWS = [
    {"Product ID": "P-200", "Description": "Wireless Mouse",            "Price": "25.50",  "Location": "Warehouse A, Shelf 2"},
    {"Product ID": "P-201", "Description": "Keyboard, Mechanical",      "Price": "89.99",  "Location": "Store Front"},
    {"Product ID": "P-202", "Description": "Monitor 24-inch",           "Price": "150.00", "Location": "Paris, France"},
    {"Product ID": "P-203", "Description": "USB Cable",                 "Price": "5.99",   "Location": "OnlineP"},
]

CSV_ROWS = [
    {"Product ID": "P-200", "Description": "Wireless Mouse",            "Price": "25.50",  "Location": "Warehouse A, Shelf 2"},
    {"Product ID": "P-201", "Description": "Keyboard, Mechanical, RGB", "Price": "89.99",  "Location": "Store Front"},
    {"Product ID": "P-202", "Description": "Monitor 24-inch",           "Price": "150.00", "Location": "Paris, France"},
    {"Product ID": "P-203", "Description": "USB Cable, 3ft",            "Price": "5.99",   "Location": "Online"},
]

XLS_FIRST_ROW = {
    "Index": "1.0", "First Name": "Dulce", "Last Name": "Abril",
    "Gender": "Female", "Country": "United States", "Age": "32.0",
}
XLS_LAST_ROW = {
    "Index": "9.0", "First Name": "Vincenza", "Last Name": "Weiland",
    "Gender": "Female", "Country": "United States", "Age": "40.0",
}


# ---------------------------------------------------------------------------
# xlsx edit
# ---------------------------------------------------------------------------

class TestGetRowsXlsxEdit:

    def test_correct_row_count(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert len(lib.get_rows()) == 4

    def test_all_rows_match_expected(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert lib.get_rows() == XLSX_ROWS

    def test_each_row_has_all_four_keys(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        for row in lib.get_rows():
            assert list(row.keys()) == XLSX_HEADERS

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

    def test_header_row_out_of_range_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(HeadersNotDeterminedException):
            lib.get_rows(header_row=9999)


# ---------------------------------------------------------------------------
# xlsx stream
# ---------------------------------------------------------------------------

class TestGetRowsXlsxStream:

    def test_correct_row_count(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        assert len(lib.get_rows()) == 4

    def test_all_rows_match_expected(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        assert lib.get_rows() == XLSX_ROWS

    def test_produces_identical_result_to_edit_mode(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        stream_rows = lib.get_rows()
        lib.close()
        lib.load_workbook(XLSX_FILE)
        assert lib.get_rows() == stream_rows

    def test_calling_get_rows_twice_raises_streaming_violation(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.get_rows()
        with pytest.raises(StreamingViolationException):
            lib.get_rows()


# ---------------------------------------------------------------------------
# xls standard
# ---------------------------------------------------------------------------

class TestGetRowsXlsStandard:

    def test_correct_row_count(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        assert len(lib.get_rows()) == 9

    def test_first_row_content(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        assert lib.get_rows()[0] == XLS_FIRST_ROW

    def test_last_row_content(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        assert lib.get_rows()[-1] == XLS_LAST_ROW

    def test_numeric_values_stringified_as_floats(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        rows = lib.get_rows()
        assert rows[0]["Index"] == "1.0"
        assert rows[0]["Age"] == "32.0"

    def test_trailing_empty_columns_excluded_from_result(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        rows = lib.get_rows()
        assert "" not in rows[0]
        assert list(rows[0].keys()) == ["Index", "First Name", "Last Name", "Gender", "Country", "Age"]

    def test_all_rows_contain_expected_name_columns(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        for row in lib.get_rows():
            assert "First Name" in row
            assert "Last Name" in row
            assert "Country" in row


# ---------------------------------------------------------------------------
# xls on demand
# ---------------------------------------------------------------------------

class TestGetRowsXlsOnDemand:

    def test_correct_row_count(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE, read_only=True)
        assert len(lib.get_rows()) == 9

    def test_produces_identical_result_to_standard_mode(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE, read_only=True)
        on_demand_rows = lib.get_rows()
        lib.close()
        lib.load_workbook(XLS_FILE)
        assert lib.get_rows() == on_demand_rows


# ---------------------------------------------------------------------------
# csv edit
# ---------------------------------------------------------------------------

class TestGetRowsCsvEdit:

    def test_correct_row_count(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        assert len(lib.get_rows()) == 4

    def test_all_rows_match_expected(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        assert lib.get_rows() == CSV_ROWS

    def test_quoted_field_with_comma_is_single_value(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        rows = lib.get_rows()
        assert rows[1]["Description"] == "Keyboard, Mechanical, RGB"
        assert rows[0]["Location"] == "Warehouse A, Shelf 2"

    def test_all_rows_have_all_four_header_keys(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        for row in lib.get_rows():
            assert list(row.keys()) == ["Product ID", "Description", "Price", "Location"]

    def test_header_row_out_of_range_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        with pytest.raises(HeadersNotDeterminedException):
            lib.get_rows(header_row=9999)


# ---------------------------------------------------------------------------
# csv stream
# ---------------------------------------------------------------------------

class TestGetRowsCsvStream:

    def test_correct_row_count(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        assert len(lib.get_rows()) == 4

    def test_all_rows_match_expected(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        assert lib.get_rows() == CSV_ROWS

    def test_produces_identical_result_to_edit_mode(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        stream_rows = lib.get_rows()
        lib.close()
        lib.load_workbook(CSV_FILE)
        assert lib.get_rows() == stream_rows

    def test_calling_get_rows_twice_raises_streaming_violation(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        lib.get_rows()
        with pytest.raises(StreamingViolationException):
            lib.get_rows()


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


    def test_exact_match_dict_single_criteria_returns_one_row(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows(search_criteria={"Product ID": "P-200"})
        assert len(rows) == 1
        assert rows[0]["Product ID"] == "P-200"

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
        rows = lib.get_rows(search_criteria="Product ID=P-200;Price=25.50")
        assert len(rows) == 1
        assert rows[0]["Product ID"] == "P-200"

    def test_string_criteria_no_criteria_returns_all(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert lib.get_rows() == lib.get_rows(search_criteria=None)


    def test_and_logic_two_criteria_narrows_result_to_one_row(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows(search_criteria={"Product ID": "P-202", "Price": "150.00"})
        assert len(rows) == 1
        assert rows[0]["Product ID"] == "P-202"

    def test_and_logic_conflicting_criteria_returns_empty(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows(search_criteria={"Product ID": "P-200", "Price": "150.00"})
        assert rows == []


    def test_partial_match_true_substring_matches(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows(search_criteria={"Description": "Keyboard"}, partial_match=True)
        assert len(rows) == 1
        assert rows[0]["Product ID"] == "P-201"

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


    def test_exact_match_on_csv(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        rows = lib.get_rows(search_criteria={"Product ID": "P-202"})
        assert len(rows) == 1
        assert rows[0]["Location"] == "Paris, France"

    def test_partial_match_on_csv_comma_in_value(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        rows = lib.get_rows(search_criteria={"Description": "Keyboard"}, partial_match=True)
        assert len(rows) == 1
        assert rows[0]["Description"] == "Keyboard, Mechanical, RGB"


    def test_exact_match_on_xls(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        rows = lib.get_rows(search_criteria={"First Name": "Dulce"})
        assert len(rows) == 1
        assert rows[0]["Last Name"] == "Abril"

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

    def test_one_row_returns_dict_not_list(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        result = lib.get_rows(one_row=True)
        assert isinstance(result, dict)
        assert not isinstance(result, list)

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

    def test_one_row_stops_after_first_match(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        result = lib.get_rows(one_row=True)
        assert result["Product ID"] == "P-200"

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

    def test_one_row_on_csv(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        result = lib.get_rows(search_criteria={"Product ID": "P-203"}, one_row=True)
        assert isinstance(result, dict)
        assert result["Description"] == "USB Cable, 3ft"

    def test_one_row_on_xls(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        result = lib.get_rows(search_criteria={"First Name": "Dulce"}, one_row=True)
        assert isinstance(result, dict)
        assert result["Last Name"] == "Abril"

    def test_one_row_no_workbook_raises(self, lib: RFExcelLibrary):
        with pytest.raises(WorkbookNotOpenException):
            lib.get_rows(one_row=True)

    def test_one_row_early_exit_does_not_exhaust_all_rows(self, lib: RFExcelLibrary, tmp_path: Path):
        lib.load_workbook(XLSX_FILE)
        first = lib.get_rows(one_row=True)
        assert first == XLSX_ROWS[0]
        all_rows = lib.get_rows()
        assert len(all_rows) == 4
