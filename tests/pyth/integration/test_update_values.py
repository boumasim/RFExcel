"""Integration tests for the Update Values keyword.

Tests verify in-memory updates only. Persistence is covered by save_workbook tests.
CSV edit tests use tmp_path copies because CsvEditResource auto-saves on close().

File layouts used:
  data.xlsx  – headers at row 1: Product ID | Description | Price | Location
               data rows 2-5: P-200…P-203
  data.csv   – headers at row 1: Product ID | Description | Price | Location
               data rows 2-5: P-200…P-203
  example.xls – headers at row 1: Index | First Name | Last Name | Gender | Country | Age

Covers:
  - XLSX edit: exact match updates; count returned.
  - XLSX edit: partial_match=True updates substring matches.
  - XLSX edit: unmatched criteria → 0 updates, data unchanged.
  - XLSX edit: only specified columns are updated; others untouched.
  - XLSX edit: multiple matching rows all updated.
  - XLSX edit: header_row out of range → HeadersNotDeterminedException.
  - XLSX edit: string search_criteria (key=value;…) accepted.
  - XLSX streaming → LibraryException.
  - XLS edit: lazy conversion triggered; in-memory update verified.
  - CSV edit: exact match, partial match.
  - CSV streaming → LibraryException.
  - No workbook open: returns 0 silently.
  - Update First: only first match updated; returns True/False; no-workbook → False.
"""
import shutil

import pytest

from rfexcel.exception.library_exceptions import (
    HeadersNotDeterminedException, LibraryException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

# xlsx has headers in row 1 (first sheet 'List 1')
_XLSX_HEADER_ROW = 1


# ---------------------------------------------------------------------------
# XLSX – Edit mode
# ---------------------------------------------------------------------------

class TestUpdateValuesXlsxEdit:

    def test_returns_count_of_updated_rows(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        count = lib.update_values(
            search_criteria={"Product ID": "P-200"},
            values={"Price": "0.00"},
            header_row=_XLSX_HEADER_ROW,
        )
        assert count == 1

    def test_matching_column_is_changed(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.update_values(
            search_criteria={"Product ID": "P-200"},
            values={"Price": "0.00"},
            header_row=_XLSX_HEADER_ROW,
        )
        rows = lib.get_rows(header_row=_XLSX_HEADER_ROW)
        row = next(r for r in rows if r["Product ID"] == "P-200")
        assert row["Price"] == "0.00"

    def test_unspecified_columns_are_not_touched(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows_before = lib.get_rows(header_row=_XLSX_HEADER_ROW)
        original_desc = next(r["Description"] for r in rows_before if r["Product ID"] == "P-201")

        lib.update_values(
            search_criteria={"Product ID": "P-201"},
            values={"Price": "1.00"},
            header_row=_XLSX_HEADER_ROW,
        )
        rows = lib.get_rows(header_row=_XLSX_HEADER_ROW)
        row = next(r for r in rows if r["Product ID"] == "P-201")
        assert row["Description"] == original_desc  # unchanged
        assert row["Price"] == "1.00"               # updated

    def test_no_match_returns_zero_and_leaves_data_unchanged(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows_before = lib.get_rows(header_row=_XLSX_HEADER_ROW)
        count = lib.update_values(
            search_criteria={"Product ID": "NONEXISTENT"},
            values={"Price": "999.99"},
            header_row=_XLSX_HEADER_ROW,
        )
        assert count == 0
        assert lib.get_rows(header_row=_XLSX_HEADER_ROW) == rows_before

    def test_multiple_matching_rows_all_updated(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.update_values(
            search_criteria={"Product ID": "P-201"},
            values={"Location": "SAME"},
            header_row=_XLSX_HEADER_ROW,
        )
        lib.update_values(
            search_criteria={"Product ID": "P-202"},
            values={"Location": "SAME"},
            header_row=_XLSX_HEADER_ROW,
        )
        count = lib.update_values(
            search_criteria={"Location": "SAME"},
            values={"Price": "0.00"},
            header_row=_XLSX_HEADER_ROW,
        )
        assert count == 2
        rows = lib.get_rows(header_row=_XLSX_HEADER_ROW)
        for r in rows:
            if r["Location"] == "SAME":
                assert r["Price"] == "0.00"

    def test_partial_match_updates_substring_rows(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        count = lib.update_values(
            search_criteria={"Location": "Warehouse"},
            values={"Price": "0.01"},
            header_row=_XLSX_HEADER_ROW,
            partial_match=True,
        )
        assert count >= 1
        rows = lib.get_rows(header_row=_XLSX_HEADER_ROW)
        updated = [r for r in rows if "Warehouse" in r["Location"]]
        assert all(r["Price"] == "0.01" for r in updated)

    def test_string_search_criteria_accepted(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        count = lib.update_values(
            search_criteria="Product ID=P-203",
            values={"Location": "Archived"},
            header_row=_XLSX_HEADER_ROW,
        )
        assert count == 1
        rows = lib.get_rows(header_row=_XLSX_HEADER_ROW)
        row = next(r for r in rows if r["Product ID"] == "P-203")
        assert row["Location"] == "Archived"

    def test_header_row_out_of_range_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(HeadersNotDeterminedException):
            lib.update_values(
                search_criteria={"Product ID": "P-200"},
                values={"Price": "0.00"},
                header_row=9999,
            )

    def test_unknown_value_keys_silently_ignored(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows_before = lib.get_rows(header_row=_XLSX_HEADER_ROW)
        lib.update_values(
            search_criteria={"Product ID": "P-200"},
            values={"NonExistentColumn": "X"},
            header_row=_XLSX_HEADER_ROW,
        )
        assert lib.get_rows(header_row=_XLSX_HEADER_ROW) == rows_before


# ---------------------------------------------------------------------------
# XLSX – Streaming mode
# ---------------------------------------------------------------------------

class TestUpdateValuesXlsxStream:

    def test_raises_in_stream_mode(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        with pytest.raises(LibraryException):
            lib.update_values(
                search_criteria={"Product ID": "P-200"},
                values={"Price": "0.00"},
                header_row=_XLSX_HEADER_ROW,
            )


# ---------------------------------------------------------------------------
# XLS – Edit mode (lazy conversion)
# ---------------------------------------------------------------------------

class TestUpdateValuesXlsEdit:

    def test_update_triggers_conversion(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        count = lib.update_values(
            search_criteria={"First Name": "Dulce"},
            values={"Country": "Updated"},
        )
        assert count == 1
        row = next(r for r in lib.get_rows() if r["First Name"] == "Dulce")
        assert row["Country"] == "Updated"


# ---------------------------------------------------------------------------
# CSV – Edit mode
# ---------------------------------------------------------------------------

class TestUpdateValuesCsvEdit:

    def test_matching_row_updated(self, lib: RFExcelLibrary, tmp_path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        count = lib.update_values(
            search_criteria={"Product ID": "P-200"},
            values={"Price": "0.00"},
        )
        assert count == 1
        row = next(r for r in lib.get_rows() if r["Product ID"] == "P-200")
        assert row["Price"] == "0.00"

    def test_unspecified_columns_untouched(self, lib: RFExcelLibrary, tmp_path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        original_desc = next(
            r["Description"] for r in lib.get_rows() if r["Product ID"] == "P-201"
        )
        lib.update_values(
            search_criteria={"Product ID": "P-201"},
            values={"Price": "1.11"},
        )
        row = next(r for r in lib.get_rows() if r["Product ID"] == "P-201")
        assert row["Description"] == original_desc
        assert row["Price"] == "1.11"

    def test_partial_match(self, lib: RFExcelLibrary, tmp_path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        count = lib.update_values(
            search_criteria={"Location": "Online"},
            values={"Price": "FREE"},
            partial_match=True,
        )
        assert count >= 1
        rows = lib.get_rows()
        assert all(
            r["Price"] == "FREE" for r in rows if "Online" in r["Location"]
        )


# ---------------------------------------------------------------------------
# CSV – Streaming mode
# ---------------------------------------------------------------------------

class TestUpdateValuesCsvStream:

    def test_raises_in_stream_mode(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        with pytest.raises(LibraryException):
            lib.update_values(
                search_criteria={"Product ID": "P-200"},
                values={"Price": "0.00"},
            )


# ---------------------------------------------------------------------------
# No workbook open
# ---------------------------------------------------------------------------

class TestUpdateValuesNoWorkbook:

    def test_returns_zero_when_no_workbook_open(self, lib: RFExcelLibrary):
        result = lib.update_values(
            search_criteria={"Product ID": "P-200"},
            values={"Price": "0.00"},
        )
        assert result == 0


# ---------------------------------------------------------------------------
# Update Values with first_only=True
# ---------------------------------------------------------------------------

class TestUpdateFirst:

    def test_returns_1_when_match_found(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        count = lib.update_values(
            search_criteria={"Product ID": "P-200"},
            values={"Price": "0.01"},
            first_only=True,
        )
        assert count == 1

    def test_returns_0_when_no_match(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        count = lib.update_values(
            search_criteria={"Product ID": "NONEXISTENT"},
            values={"Price": "0.01"},
            first_only=True,
        )
        assert count == 0

    def test_only_first_match_is_updated(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        # Set two rows to the same location so both would match
        lib.update_values(search_criteria={"Product ID": "P-201"}, values={"Location": "SHARED"})
        lib.update_values(search_criteria={"Product ID": "P-202"}, values={"Location": "SHARED"})

        lib.update_values(
            search_criteria={"Location": "SHARED"},
            values={"Price": "FIRST_ONLY"},
            first_only=True,
        )
        rows = lib.get_rows()
        updated = [r for r in rows if r["Location"] == "SHARED"]
        assert len(updated) == 2
        assert sum(1 for r in updated if r["Price"] == "FIRST_ONLY") == 1

    def test_correct_column_changed(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.update_values(
            search_criteria={"Product ID": "P-203"},
            values={"Location": "Updated"},
            first_only=True,
        )
        rows = lib.get_rows()
        row = next(r for r in rows if r["Product ID"] == "P-203")
        assert row["Location"] == "Updated"

    def test_returns_0_when_no_workbook_open(self, lib: RFExcelLibrary):
        assert lib.update_values(
            search_criteria={"Product ID": "P-200"},
            values={"Price": "0.00"},
            first_only=True,
        ) == 0

    def test_csv_only_first_match_updated(self, lib: RFExcelLibrary, tmp_path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        count = lib.update_values(
            search_criteria={"Location": "Online"},
            values={"Price": "FIRST"},
            partial_match=True,
            first_only=True,
        )
        assert count == 1
        rows = lib.get_rows()
        updated = [r for r in rows if r["Price"] == "FIRST"]
        assert len(updated) == 1
