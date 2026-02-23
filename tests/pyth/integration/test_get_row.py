"""Integration tests for the Get Row keyword.

Covers all six format/mode combinations and both return modes (list vs dict).

data.xlsx / data.csv (5 total rows: 1 header + 4 data):
  Row 1 (header) : Product ID | Description | Price | Location
  Row 2          : P-200 | Wireless Mouse | 25.50 | Warehouse A, Shelf 2
  Row 3          : P-201 | Keyboard, Mechanical[, RGB in csv] | 89.99 | Store Front
  Row 4          : P-202 | Monitor 24-inch | 150.00 | Paris, France
  Row 5          : P-203 | USB Cable[, 3ft in csv] | 5.99 | Online[P in xlsx]

example.xls (10 total rows: 1 header + 9 data, 8 physical columns):
  Row 1 (header) : Index | First Name | Last Name | Gender | Country | Age | '' | ''
  Row 2          : 1.0 | Dulce | Abril | Female | United States | 32.0 | '' | ''
  Row 10         : 9.0 | Vincenza | Weiland | Female | United States | 40.0 | '' | ''
"""
import pytest

from rfexcel.exception.library_exceptions import (FileDoesNotExistException,
                                                  StreamingViolationException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

XLSX_HEADERS = ["Product ID", "Description", "Price", "Location"]
XLS_HEADERS  = ["Index", "First Name", "Last Name", "Gender", "Country", "Age", "", ""]

# Expected list-form rows (all values are strings)
XLSX_ROW2_LIST = ["P-200", "Wireless Mouse",            "25.50",  "Warehouse A, Shelf 2"]
XLSX_ROW3_LIST = ["P-201", "Keyboard, Mechanical",      "89.99",  "Store Front"]
XLSX_ROW5_LIST = ["P-203", "USB Cable",                 "5.99",   "OnlineP"]

CSV_ROW2_LIST  = ["P-200", "Wireless Mouse",            "25.50",  "Warehouse A, Shelf 2"]
CSV_ROW3_LIST  = ["P-201", "Keyboard, Mechanical, RGB", "89.99",  "Store Front"]

XLS_ROW2_LIST  = ["1.0", "Dulce", "Abril", "Female", "United States", "32.0", "", ""]
XLS_ROW10_LIST = ["9.0", "Vincenza", "Weiland", "Female", "United States", "40.0", "", ""]

# Expected dict-form rows (mapped against XLSX_HEADERS / XLS_HEADERS)
XLSX_ROW2_DICT = {"Product ID": "P-200", "Description": "Wireless Mouse",            "Price": "25.50",  "Location": "Warehouse A, Shelf 2"}
XLSX_ROW5_DICT = {"Product ID": "P-203", "Description": "USB Cable",                 "Price": "5.99",   "Location": "OnlineP"}

CSV_ROW2_DICT  = {"Product ID": "P-200", "Description": "Wireless Mouse",            "Price": "25.50",  "Location": "Warehouse A, Shelf 2"}
CSV_ROW3_DICT  = {"Product ID": "P-201", "Description": "Keyboard, Mechanical, RGB", "Price": "89.99",  "Location": "Store Front"}

XLS_ROW2_DICT  = {"Index": "1.0", "First Name": "Dulce",    "Last Name": "Abril",   "Gender": "Female", "Country": "United States", "Age": "32.0", "": ""}
XLS_ROW10_DICT = {"Index": "9.0", "First Name": "Vincenza", "Last Name": "Weiland", "Gender": "Female", "Country": "United States", "Age": "40.0", "": ""}


# ─── xlsx edit mode ─────────────────────────────────────────────────────────────

class TestGetRowXlsxEdit:

    def test_row_without_headers_returns_list(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert isinstance(lib.get_row(2), list)

    def test_row_with_headers_returns_dict(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert isinstance(lib.get_row(2, headers=XLSX_HEADERS), dict)

    def test_row2_list_values(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert lib.get_row(2) == XLSX_ROW2_LIST

    def test_row3_list_values(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert lib.get_row(3) == XLSX_ROW3_LIST

    def test_row5_list_values(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert lib.get_row(5) == XLSX_ROW5_LIST

    def test_row2_dict_values(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert lib.get_row(2, headers=XLSX_HEADERS) == XLSX_ROW2_DICT

    def test_row5_dict_values(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert lib.get_row(5, headers=XLSX_HEADERS) == XLSX_ROW5_DICT

    def test_list_length_equals_column_count(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert len(lib.get_row(2)) == 4

    def test_header_row_itself_accessible(self, lib: RFExcelLibrary):
        """Row 1 is the header row — it must be readable as a plain list."""
        lib.load_workbook(XLSX_FILE)
        assert lib.get_row(1) == XLSX_HEADERS

    def test_out_of_bounds_row_returns_empty_list(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert lib.get_row(99) == []

    def test_repeated_calls_return_same_row(self, lib: RFExcelLibrary):
        """Edit mode supports random access — same row twice must be equal."""
        lib.load_workbook(XLSX_FILE)
        assert lib.get_row(2) == lib.get_row(2)

    def test_dict_keys_match_supplied_headers(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        row = lib.get_row(2, headers=XLSX_HEADERS)
        assert isinstance(row, dict)
        assert list(row.keys()) == XLSX_HEADERS


# ─── xlsx stream mode ───────────────────────────────────────────────────────────

class TestGetRowXlsxStream:

    def test_row_without_headers_returns_list(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        assert isinstance(lib.get_row(1), list)

    def test_first_row_is_header_row(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        assert lib.get_row(1) == XLSX_HEADERS

    def test_sequential_rows_are_consistent(self, lib: RFExcelLibrary):
        """Stream mode must yield rows in order."""
        lib.load_workbook(XLSX_FILE, read_only=True)
        row1 = lib.get_row(1)
        row2 = lib.get_row(2)
        assert row1 == XLSX_HEADERS
        assert row2 == XLSX_ROW2_LIST

    def test_row_with_headers_returns_dict(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.get_row(1)  # consume header row first
        assert isinstance(lib.get_row(2, headers=XLSX_HEADERS), dict)

    def test_stream_row2_dict_matches_edit_mode(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.get_row(1)  # consume header row first
        assert lib.get_row(2, headers=XLSX_HEADERS) == XLSX_ROW2_DICT

    def test_re_reading_same_row_raises_streaming_violation(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.get_row(1)
        with pytest.raises(StreamingViolationException):
            lib.get_row(1)

    def test_reading_earlier_row_raises_streaming_violation(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.get_row(1)
        lib.get_row(2)
        with pytest.raises(StreamingViolationException):
            lib.get_row(1)


# ─── xls standard (edit) mode ───────────────────────────────────────────────────

class TestGetRowXlsStandard:

    def test_row_without_headers_returns_list(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        assert isinstance(lib.get_row(2), list)

    def test_row2_list_values(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        assert lib.get_row(2) == XLS_ROW2_LIST

    def test_row10_list_values(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        assert lib.get_row(10) == XLS_ROW10_LIST

    def test_row2_dict_values(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        assert lib.get_row(2, headers=XLS_HEADERS) == XLS_ROW2_DICT

    def test_row10_dict_values(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        assert lib.get_row(10, headers=XLS_HEADERS) == XLS_ROW10_DICT

    def test_numeric_values_stringified_as_floats(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        row = lib.get_row(2)
        assert row[0] == "1.0"   # Index
        assert row[5] == "32.0"  # Age

    def test_trailing_empty_columns_present_in_list(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        row = lib.get_row(2)
        assert row[6] == ""
        assert row[7] == ""

    def test_repeated_calls_return_same_row(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        assert lib.get_row(2) == lib.get_row(2)

    def test_out_of_bounds_row_returns_empty_list(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        assert lib.get_row(99) == []


# ─── xls on_demand mode ─────────────────────────────────────────────────────────

class TestGetRowXlsOnDemand:

    def test_row2_list_matches_standard_mode(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE, read_only=True)
        on_demand = lib.get_row(2)
        lib.close()
        lib.load_workbook(XLS_FILE)
        assert lib.get_row(2) == on_demand

    def test_row2_dict_values(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE, read_only=True)
        assert lib.get_row(2, headers=XLS_HEADERS) == XLS_ROW2_DICT

    def test_random_access_row10_then_row2(self, lib: RFExcelLibrary):
        """xls on_demand is random-access, not streaming — both orders must work."""
        lib.load_workbook(XLS_FILE, read_only=True)
        row10 = lib.get_row(10)
        row2  = lib.get_row(2)
        assert row10 == XLS_ROW10_LIST
        assert row2  == XLS_ROW2_LIST


# ─── csv edit mode ───────────────────────────────────────────────────────────────

class TestGetRowCsvEdit:

    def test_row_without_headers_returns_list(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        assert isinstance(lib.get_row(2), list)

    def test_row2_list_values(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        assert lib.get_row(2) == CSV_ROW2_LIST

    def test_row3_list_values(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        assert lib.get_row(3) == CSV_ROW3_LIST

    def test_row2_dict_values(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        assert lib.get_row(2, headers=XLSX_HEADERS) == CSV_ROW2_DICT

    def test_row3_dict_values(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        assert lib.get_row(3, headers=XLSX_HEADERS) == CSV_ROW3_DICT

    def test_quoted_field_with_comma_is_single_value(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        row = lib.get_row(3)
        assert row[1] == "Keyboard, Mechanical, RGB"

    def test_header_row_readable_as_list(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        assert lib.get_row(1) == XLSX_HEADERS

    def test_repeated_calls_return_same_row(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        assert lib.get_row(2) == lib.get_row(2)

    def test_out_of_bounds_row_returns_empty_list(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        assert lib.get_row(99) == []


# ─── csv stream mode ─────────────────────────────────────────────────────────────

class TestGetRowCsvStream:

    def test_first_row_is_header_row(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        assert lib.get_row(1) == XLSX_HEADERS

    def test_sequential_rows_are_consistent(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        row1 = lib.get_row(1)
        row2 = lib.get_row(2)
        assert row1 == XLSX_HEADERS
        assert row2 == CSV_ROW2_LIST

    def test_row_with_headers_returns_dict(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        lib.get_row(1)  # consume header row first
        assert isinstance(lib.get_row(2, headers=XLSX_HEADERS), dict)

    def test_stream_row2_dict_matches_edit_mode(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        lib.get_row(1)  # consume header row first
        assert lib.get_row(2, headers=XLSX_HEADERS) == CSV_ROW2_DICT

    def test_re_reading_same_row_raises_streaming_violation(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        lib.get_row(1)
        with pytest.raises(StreamingViolationException):
            lib.get_row(1)

    def test_reading_earlier_row_raises_streaming_violation(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        lib.get_row(1)
        lib.get_row(2)
        with pytest.raises(StreamingViolationException):
            lib.get_row(1)


# ─── negative / edge ─────────────────────────────────────────────────────────────

class TestGetRowNegative:

    def test_returns_empty_list_when_no_workbook_loaded(self, lib: RFExcelLibrary):
        assert lib.get_row(1) == []

    def test_returns_empty_list_after_close(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.close()
        assert lib.get_row(2) == []

    def test_partial_headers_list_maps_as_many_as_provided(self, lib: RFExcelLibrary):
        """Fewer headers than columns — extra columns get fillvalue ''."""
        lib.load_workbook(XLSX_FILE)
        row = lib.get_row(2, headers=["Product ID", "Description"])
        assert isinstance(row, dict)
        assert row["Product ID"] == "P-200"
        assert row["Description"] == "Wireless Mouse"

    def test_empty_headers_list_returns_list_not_dict(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        result = lib.get_row(2, headers=[])
        assert isinstance(result, list)
