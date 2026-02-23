"""Integration tests for the Load Workbook keyword.

Covers:
  - All six format/mode combinations (xlsx edit, xlsx stream, xls standard,
    xls on_demand, csv edit, csv stream).
  - Verifying the workbook is actually usable after loading (Get Rows sanity check).
  - Negative: non-existent file, unsupported extension.
  - Edge: loading a second file after the first is still open (replaces it).
"""
import pytest

from rfexcel.exception.library_exceptions import (
    FileDoesNotExistException, FileFormatNotSupportedException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

# ─── positive ─────────────────────────────────────────────────────────────────

class TestLoadWorkbookPositive:

    def test_load_xlsx_edit_mode_sets_active_workbook(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert lib._active_workbook is not None

    def test_load_xlsx_stream_mode_sets_active_workbook(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        assert lib._active_workbook is not None

    def test_load_xls_edit_mode_sets_active_workbook(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        assert lib._active_workbook is not None

    def test_load_xls_on_demand_mode_sets_active_workbook(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE, read_only=True)
        assert lib._active_workbook is not None

    def test_load_csv_edit_mode_sets_active_workbook(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        assert lib._active_workbook is not None

    def test_load_csv_stream_mode_sets_active_workbook(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        assert lib._active_workbook is not None

    def test_load_xlsx_edit_is_immediately_readable(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows()
        assert len(rows) == 4
        assert rows[0]["Product ID"] == "P-200"

    def test_load_xlsx_stream_is_immediately_readable(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        rows = lib.get_rows()
        assert len(rows) == 4

    def test_load_xls_edit_is_immediately_readable(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        rows = lib.get_rows()
        assert len(rows) == 9

    def test_load_xls_on_demand_is_immediately_readable(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE, read_only=True)
        rows = lib.get_rows()
        assert len(rows) == 9

    def test_load_csv_edit_is_immediately_readable(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        rows = lib.get_rows()
        assert len(rows) == 4

    def test_load_csv_stream_is_immediately_readable(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        rows = lib.get_rows()
        assert len(rows) == 4


# ─── negative ─────────────────────────────────────────────────────────────────

class TestLoadWorkbookNegative:

    def test_non_existent_xlsx_raises(self, lib: RFExcelLibrary):
        with pytest.raises(FileDoesNotExistException):
            lib.load_workbook("/nonexistent/path/missing.xlsx")

    def test_non_existent_csv_raises(self, lib: RFExcelLibrary):
        with pytest.raises(FileDoesNotExistException):
            lib.load_workbook("/nonexistent/path/missing.csv")

    def test_non_existent_xls_raises(self, lib: RFExcelLibrary):
        with pytest.raises(FileDoesNotExistException):
            lib.load_workbook("/nonexistent/path/missing.xls")

    def test_unsupported_extension_raises(self, lib: RFExcelLibrary):
        with pytest.raises(FileFormatNotSupportedException):
            lib.load_workbook("/some/path/file.txt")

    def test_unsupported_ods_extension_raises(self, lib: RFExcelLibrary):
        with pytest.raises(FileFormatNotSupportedException):
            lib.load_workbook("/some/path/file.ods")

    def test_active_workbook_is_none_after_failed_load(self, lib: RFExcelLibrary):
        with pytest.raises(FileDoesNotExistException):
            lib.load_workbook("/nonexistent/path/missing.xlsx")
        assert lib._active_workbook is None


# ─── edge cases ───────────────────────────────────────────────────────────────

class TestLoadWorkbookEdge:

    def test_loading_second_file_replaces_first(self, lib: RFExcelLibrary):
        """Loading a new file while one is open must replace the active workbook."""
        lib.load_workbook(XLSX_FILE)
        first_wb = lib._active_workbook
        lib.load_workbook(CSV_FILE)
        assert lib._active_workbook is not first_wb

    def test_loading_after_close_works(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.close()
        lib.load_workbook(XLSX_FILE)
        assert lib._active_workbook is not None
        assert len(lib.get_rows()) == 4

    def test_xlsx_edit_and_stream_produce_identical_rows(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        edit_rows = lib.get_rows()
        lib.close()
        lib.load_workbook(XLSX_FILE, read_only=True)
        stream_rows = lib.get_rows()
        assert edit_rows == stream_rows

    def test_csv_edit_and_stream_produce_identical_rows(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        edit_rows = lib.get_rows()
        lib.close()
        lib.load_workbook(CSV_FILE, read_only=True)
        stream_rows = lib.get_rows()
        assert edit_rows == stream_rows

    def test_xls_edit_and_on_demand_produce_identical_rows(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        edit_rows = lib.get_rows()
        lib.close()
        lib.load_workbook(XLS_FILE, read_only=True)
        on_demand_rows = lib.get_rows()
        assert edit_rows == on_demand_rows
