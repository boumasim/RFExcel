"""Integration tests for the Delete Rows keyword.

CSV edit tests use tmp_path copies because CsvEditResource auto-saves on close().

File layouts used:
  data.xlsx  – headers at row 1: Product ID | Description | Price | Location
               data rows 2-5: P-200…P-203 (4 rows total)
  data.csv   – headers at row 1: same layout
  example.xls – headers at row 1: Index | First Name | Last Name | Gender | Country | Age

Covers:
  - XLSX edit: matched row is removed; row count decreases; count returned.
  - XLSX edit: first_only=True removes only one row when multiple match.
  - XLSX edit: no match → 0 deleted, row count unchanged.
  - XLSX edit: partial_match=True deletes substring matches.
  - XLSX edit: remaining rows still readable after deletion.
  - XLSX edit: header_row out of range → HeadersNotDeterminedException.
  - XLSX streaming → LibraryException.
  - XLS edit: lazy conversion triggered; row deleted in memory.
  - CSV edit: matched row deleted; row count decreases.
  - CSV edit: first_only=True removes only first match.
  - CSV streaming → LibraryException.
  - No workbook open: returns 0 silently.
"""
import shutil

import pytest

from rfexcel.exception.library_exceptions import (
    HeadersNotDeterminedException, LibraryException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

# ---------------------------------------------------------------------------
# XLSX – Edit mode
# ---------------------------------------------------------------------------

class TestDeleteRowsXlsxEdit:

    def test_returns_count_of_deleted_rows(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        count = lib.delete_rows(search_criteria={"Product ID": "P-200"})
        assert count == 1

    def test_row_is_removed(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.delete_rows(search_criteria={"Product ID": "P-200"})
        rows = lib.get_rows()
        assert all(r["Product ID"] != "P-200" for r in rows)

    def test_row_count_decreases(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        before = len(lib.get_rows())
        lib.delete_rows(search_criteria={"Product ID": "P-201"})
        assert len(lib.get_rows()) == before - 1

    def test_no_match_returns_zero(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        before = lib.get_rows()
        count = lib.delete_rows(search_criteria={"Product ID": "NONEXISTENT"})
        assert count == 0
        assert lib.get_rows() == before

    def test_multiple_matches_all_deleted(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        # Tag two rows with the same location first
        lib.update_values(search_criteria={"Product ID": "P-201"}, values={"Location": "SAME"})
        lib.update_values(search_criteria={"Product ID": "P-202"}, values={"Location": "SAME"})
        count = lib.delete_rows(search_criteria={"Location": "SAME"})
        assert count == 2
        rows = lib.get_rows()
        assert all(r["Location"] != "SAME" for r in rows)

    def test_first_only_deletes_single_row(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.update_values(search_criteria={"Product ID": "P-201"}, values={"Location": "SAME"})
        lib.update_values(search_criteria={"Product ID": "P-202"}, values={"Location": "SAME"})
        count = lib.delete_rows(search_criteria={"Location": "SAME"}, first_only=True)
        assert count == 1
        rows = lib.get_rows()
        # One row with SAME should still remain
        assert sum(1 for r in rows if r["Location"] == "SAME") == 1

    def test_partial_match_deletes_substring_rows(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        before = len(lib.get_rows())
        count = lib.delete_rows(search_criteria={"Location": "Warehouse"}, partial_match=True)
        assert count >= 1
        assert len(lib.get_rows()) == before - count
        assert all("Warehouse" not in r["Location"] for r in lib.get_rows())

    def test_remaining_rows_are_readable(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.delete_rows(search_criteria={"Product ID": "P-202"})
        ids = [r["Product ID"] for r in lib.get_rows()]
        assert "P-200" in ids
        assert "P-201" in ids
        assert "P-202" not in ids
        assert "P-203" in ids

    def test_header_row_out_of_range_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(HeadersNotDeterminedException):
            lib.delete_rows(search_criteria={"Product ID": "P-200"}, header_row=9999)


# ---------------------------------------------------------------------------
# XLSX – Streaming mode
# ---------------------------------------------------------------------------

class TestDeleteRowsXlsxStream:

    def test_raises_in_stream_mode(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        with pytest.raises(LibraryException):
            lib.delete_rows(search_criteria={"Product ID": "P-200"})


# ---------------------------------------------------------------------------
# XLS – Edit mode (lazy conversion)
# ---------------------------------------------------------------------------

class TestDeleteRowsXlsEdit:

    def test_delete_triggers_conversion(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        before = len(lib.get_rows())
        count = lib.delete_rows(search_criteria={"First Name": "Dulce"})
        assert count == 1
        assert len(lib.get_rows()) == before - 1
        assert all(r["First Name"] != "Dulce" for r in lib.get_rows())


# ---------------------------------------------------------------------------
# CSV – Edit mode
# ---------------------------------------------------------------------------

class TestDeleteRowsCsvEdit:

    def test_matched_row_deleted(self, lib: RFExcelLibrary, tmp_path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        count = lib.delete_rows(search_criteria={"Product ID": "P-200"})
        assert count == 1
        assert all(r["Product ID"] != "P-200" for r in lib.get_rows())

    def test_row_count_decreases(self, lib: RFExcelLibrary, tmp_path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        before = len(lib.get_rows())
        lib.delete_rows(search_criteria={"Product ID": "P-201"})
        assert len(lib.get_rows()) == before - 1

    def test_first_only_deletes_single_row(self, lib: RFExcelLibrary, tmp_path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        # Tag two rows with the same Location so first_only makes a difference
        lib.update_values(search_criteria={"Product ID": "P-200"}, values={"Location": "DUPLICATE_LOC"})
        lib.update_values(search_criteria={"Product ID": "P-201"}, values={"Location": "DUPLICATE_LOC"})
        count = lib.delete_rows(search_criteria={"Location": "DUPLICATE_LOC"}, first_only=True)
        assert count == 1
        # One row with DUPLICATE_LOC must still remain
        rows = lib.get_rows()
        assert sum(1 for r in rows if r["Location"] == "DUPLICATE_LOC") == 1


# ---------------------------------------------------------------------------
# CSV – Streaming mode
# ---------------------------------------------------------------------------

class TestDeleteRowsCsvStream:

    def test_raises_in_stream_mode(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        with pytest.raises(LibraryException):
            lib.delete_rows(search_criteria={"Product ID": "P-200"})


# ---------------------------------------------------------------------------
# No workbook open
# ---------------------------------------------------------------------------

class TestDeleteRowsNoWorkbook:

    def test_returns_zero_when_no_workbook_open(self, lib: RFExcelLibrary):
        assert lib.delete_rows(search_criteria={"Product ID": "P-200"}) == 0
