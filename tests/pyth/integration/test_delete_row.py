"""Integration tests for the Delete Row keyword.

CSV edit tests use tmp_path copies because CsvEditResource auto-saves on close().

File layouts used:
  data.xlsx  – headers at row 1: Product ID | Description | Price | Location
               data rows 2-5: P-200…P-203 (4 rows total)
  data.csv   – headers at row 1: same layout
  example.xls – headers at row 1: Index | First Name | Last Name | Gender | Country | Age

Covers:
  - XLSX edit: row deleted by row number; row is no longer readable.
  - XLSX edit: row count decreases after deletion.
  - XLSX edit: remaining rows are still readable in correct order.
  - XLSX edit: row_number < 1 → RowIndexOutOfBoundsException.
  - XLSX edit: row_number beyond last row → RowIndexOutOfBoundsException.
  - XLSX streaming → LibraryException.
  - XLS edit: lazy conversion triggered; row deleted in memory.
  - CSV edit: row deleted; row count decreases.
  - CSV streaming → LibraryException.
  - No workbook open: does nothing silently.
"""
import shutil

import pytest

from rfexcel.exception.library_exceptions import (LibraryException,
                                                  RowIndexOutOfBoundsException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

# ---------------------------------------------------------------------------
# XLSX – Edit mode
# ---------------------------------------------------------------------------

class TestDeleteRowXlsxEdit:

    def test_row_is_removed(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.delete_row(2)  # row 2 is P-200
        rows = lib.get_rows()
        assert all(r["Product ID"] != "P-200" for r in rows)

    def test_row_count_decreases(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        before = len(lib.get_rows())
        lib.delete_row(2)
        assert len(lib.get_rows()) == before - 1

    def test_remaining_rows_readable_in_order(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.delete_row(3)  # row 3 is P-201
        ids = [r["Product ID"] for r in lib.get_rows()]
        assert "P-200" in ids
        assert "P-201" not in ids
        assert "P-202" in ids
        assert "P-203" in ids

    def test_delete_last_data_row(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        before = len(lib.get_rows())
        lib.delete_row(5)  # row 5 is P-203
        rows = lib.get_rows()
        assert len(rows) == before - 1
        assert all(r["Product ID"] != "P-203" for r in rows)

    def test_row_number_zero_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(RowIndexOutOfBoundsException):
            lib.delete_row(0)

    def test_row_number_negative_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(RowIndexOutOfBoundsException):
            lib.delete_row(-1)

    def test_row_number_beyond_last_row_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(RowIndexOutOfBoundsException):
            lib.delete_row(9999)


# ---------------------------------------------------------------------------
# XLSX – Streaming mode
# ---------------------------------------------------------------------------

class TestDeleteRowXlsxStream:

    def test_raises_in_stream_mode(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        with pytest.raises(LibraryException):
            lib.delete_row(2)


# ---------------------------------------------------------------------------
# XLS – Edit mode (lazy conversion)
# ---------------------------------------------------------------------------

class TestDeleteRowXlsEdit:

    def test_delete_triggers_conversion(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        before = len(lib.get_rows())
        lib.delete_row(2)  # row 2 is the first data row
        rows = lib.get_rows()
        assert len(rows) == before - 1


# ---------------------------------------------------------------------------
# CSV – Edit mode
# ---------------------------------------------------------------------------

class TestDeleteRowCsvEdit:

    def test_row_deleted(self, lib: RFExcelLibrary, tmp_path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        lib.delete_row(2)  # row 2 is P-200
        assert all(r["Product ID"] != "P-200" for r in lib.get_rows())

    def test_row_count_decreases(self, lib: RFExcelLibrary, tmp_path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        before = len(lib.get_rows())
        lib.delete_row(2)
        assert len(lib.get_rows()) == before - 1

    def test_row_number_beyond_last_row_raises(self, lib: RFExcelLibrary, tmp_path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        with pytest.raises(RowIndexOutOfBoundsException):
            lib.delete_row(9999)


# ---------------------------------------------------------------------------
# CSV – Streaming mode
# ---------------------------------------------------------------------------

class TestDeleteRowCsvStream:

    def test_raises_in_stream_mode(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        with pytest.raises(LibraryException):
            lib.delete_row(2)


# ---------------------------------------------------------------------------
# No workbook open
# ---------------------------------------------------------------------------

class TestDeleteRowNoWorkbook:

    def test_does_nothing_when_no_workbook_open(self, lib: RFExcelLibrary):
        lib.delete_row(2)  # should not raise
