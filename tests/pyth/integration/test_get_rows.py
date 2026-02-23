"""Integration tests for the Get Rows keyword.

Expected data is derived from the actual files:

data.xlsx / data.csv (4 data rows, header on row 1):
  Headers : Product ID | Description | Price | Location

example.xls (9 data rows, header on row 1):
  Headers : Index | First Name | Last Name | Gender | Country | Age | '' | ''
  (The file has 8 physical columns; the last two are empty, so their header
   key is an empty string '')
"""
import csv

import pytest

from rfexcel.exception.library_exceptions import (FileDoesNotExistException,
                                                  StreamingViolationException)
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

# ─── expected data ─────────────────────────────────────────────────────────────

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

# example.xls has 8 physical columns; the last two headers are empty strings.
XLS_FIRST_ROW = {
    "Index": "1.0", "First Name": "Dulce", "Last Name": "Abril",
    "Gender": "Female", "Country": "United States", "Age": "32.0", "": "",
}
XLS_LAST_ROW = {
    "Index": "9.0", "First Name": "Vincenza", "Last Name": "Weiland",
    "Gender": "Female", "Country": "United States", "Age": "40.0", "": "",
}


# ─── xlsx edit mode ─────────────────────────────────────────────────────────────

class TestGetRowsXlsxEdit:

    def test_correct_row_count(self, lib):
        lib.load_workbook(XLSX_FILE)
        assert len(lib.get_rows()) == 4

    def test_all_rows_match_expected(self, lib):
        lib.load_workbook(XLSX_FILE)
        assert lib.get_rows() == XLSX_ROWS

    def test_each_row_has_all_four_keys(self, lib):
        lib.load_workbook(XLSX_FILE)
        for row in lib.get_rows():
            assert list(row.keys()) == XLSX_HEADERS

    def test_cell_containing_comma_is_not_split(self, lib):
        """'Warehouse A, Shelf 2' and 'Paris, France' contain commas."""
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows()
        assert rows[0]["Location"] == "Warehouse A, Shelf 2"
        assert rows[2]["Location"] == "Paris, France"

    def test_default_header_row_equals_explicit_header_row_1(self, lib):
        lib.load_workbook(XLSX_FILE)
        assert lib.get_rows() == lib.get_rows(header_row=1)

    def test_header_row_2_shifts_data_by_one(self, lib):
        """When row 2 is treated as headers, only 3 data rows remain."""
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows(header_row=2)
        assert len(rows) == 3
        # The key for the first column is now the value of row 2
        assert "P-200" in rows[0]

    def test_header_row_beyond_data_returns_empty_list(self, lib):
        """A header_row past the last row means no data rows exist."""
        lib.load_workbook(XLSX_FILE)
        assert lib.get_rows(header_row=5) == []


# ─── xlsx stream mode ───────────────────────────────────────────────────────────

class TestGetRowsXlsxStream:

    def test_correct_row_count(self, lib):
        lib.load_workbook(XLSX_FILE, read_only=True)
        assert len(lib.get_rows()) == 4

    def test_all_rows_match_expected(self, lib):
        lib.load_workbook(XLSX_FILE, read_only=True)
        assert lib.get_rows() == XLSX_ROWS

    def test_produces_identical_result_to_edit_mode(self, lib):
        lib.load_workbook(XLSX_FILE, read_only=True)
        stream_rows = lib.get_rows()
        lib.close()
        lib.load_workbook(XLSX_FILE)
        assert lib.get_rows() == stream_rows

    def test_calling_get_rows_twice_raises_streaming_violation(self, lib):
        """Stream is exhausted after the first call — a second call must raise."""
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.get_rows()
        with pytest.raises(StreamingViolationException):
            lib.get_rows()


# ─── xls standard (edit) mode ───────────────────────────────────────────────────

class TestGetRowsXlsStandard:

    def test_correct_row_count(self, lib):
        lib.load_workbook(XLS_FILE)
        assert len(lib.get_rows()) == 9

    def test_first_row_content(self, lib):
        lib.load_workbook(XLS_FILE)
        assert lib.get_rows()[0] == XLS_FIRST_ROW

    def test_last_row_content(self, lib):
        lib.load_workbook(XLS_FILE)
        assert lib.get_rows()[-1] == XLS_LAST_ROW

    def test_numeric_values_stringified_as_floats(self, lib):
        """xlrd returns numbers as Python floats; library must stringify them."""
        lib.load_workbook(XLS_FILE)
        rows = lib.get_rows()
        assert rows[0]["Index"] == "1.0"
        assert rows[0]["Age"] == "32.0"

    def test_trailing_empty_columns_produce_empty_string_key(self, lib):
        """example.xls has 2 trailing empty columns — their header is ''."""
        lib.load_workbook(XLS_FILE)
        rows = lib.get_rows()
        assert "" in rows[0]
        assert rows[0][""] == ""

    def test_all_rows_contain_expected_name_columns(self, lib):
        lib.load_workbook(XLS_FILE)
        for row in lib.get_rows():
            assert "First Name" in row
            assert "Last Name" in row
            assert "Country" in row


# ─── xls on_demand mode ─────────────────────────────────────────────────────────

class TestGetRowsXlsOnDemand:

    def test_correct_row_count(self, lib):
        lib.load_workbook(XLS_FILE, read_only=True)
        assert len(lib.get_rows()) == 9

    def test_produces_identical_result_to_standard_mode(self, lib):
        lib.load_workbook(XLS_FILE, read_only=True)
        on_demand_rows = lib.get_rows()
        lib.close()
        lib.load_workbook(XLS_FILE)
        assert lib.get_rows() == on_demand_rows


# ─── csv edit mode ───────────────────────────────────────────────────────────────

class TestGetRowsCsvEdit:

    def test_correct_row_count(self, lib):
        lib.load_workbook(CSV_FILE)
        assert len(lib.get_rows()) == 4

    def test_all_rows_match_expected(self, lib):
        lib.load_workbook(CSV_FILE)
        assert lib.get_rows() == CSV_ROWS

    def test_quoted_field_with_comma_is_single_value(self, lib):
        """'Keyboard, Mechanical, RGB' is quoted in CSV and must not be split."""
        lib.load_workbook(CSV_FILE)
        rows = lib.get_rows()
        assert rows[1]["Description"] == "Keyboard, Mechanical, RGB"
        assert rows[0]["Location"] == "Warehouse A, Shelf 2"

    def test_all_rows_have_all_four_header_keys(self, lib):
        lib.load_workbook(CSV_FILE)
        for row in lib.get_rows():
            assert list(row.keys()) == ["Product ID", "Description", "Price", "Location"]


# ─── csv stream mode ─────────────────────────────────────────────────────────────

class TestGetRowsCsvStream:

    def test_correct_row_count(self, lib):
        lib.load_workbook(CSV_FILE, read_only=True)
        assert len(lib.get_rows()) == 4

    def test_all_rows_match_expected(self, lib):
        lib.load_workbook(CSV_FILE, read_only=True)
        assert lib.get_rows() == CSV_ROWS

    def test_produces_identical_result_to_edit_mode(self, lib):
        lib.load_workbook(CSV_FILE, read_only=True)
        stream_rows = lib.get_rows()
        lib.close()
        lib.load_workbook(CSV_FILE)
        assert lib.get_rows() == stream_rows

    def test_calling_get_rows_twice_raises_streaming_violation(self, lib):
        """Stream is exhausted after the first call — a second call must raise."""
        lib.load_workbook(CSV_FILE, read_only=True)
        lib.get_rows()
        with pytest.raises(StreamingViolationException):
            lib.get_rows()


# ─── negative / edge ─────────────────────────────────────────────────────────────

class TestGetRowsNegative:

    def test_returns_empty_list_when_no_workbook_loaded(self, lib):
        assert lib.get_rows() == []

    def test_returns_empty_list_after_close(self, lib):
        lib.load_workbook(XLSX_FILE)
        lib.close()
        assert lib.get_rows() == []

    def test_load_nonexistent_file_raises(self, lib):
        with pytest.raises(FileDoesNotExistException):
            lib.load_workbook("/nonexistent/path/missing.xlsx")

    def test_get_rows_on_empty_created_xlsx_returns_empty_list(self, lib, tmp_path):
        lib.create_workbook(str(tmp_path / "empty.xlsx"))
        assert lib.get_rows() == []

    def test_get_rows_on_empty_created_csv_returns_empty_list(self, lib, tmp_path):
        lib.create_workbook(str(tmp_path / "empty.csv"))
        assert lib.get_rows() == []

    def test_header_row_1_on_single_row_file_returns_empty_list(self, lib, tmp_path):
        """A file with only a header row and no data rows must yield []."""
        path = tmp_path / "headers_only.csv"
        with open(path, "w", newline="") as f:
            csv.writer(f).writerow(["A", "B", "C"])
        lib.load_workbook(str(path))
        assert lib.get_rows() == []
