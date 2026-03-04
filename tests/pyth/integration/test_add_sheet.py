"""Integration tests for the Add Sheet feature."""
import pytest

from rfexcel.exception.library_exceptions import (
    LibraryException, OperationNotSupportedForFormat)
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

# ---------------------------------------------------------------------------
# XLSX – Edit mode
# ---------------------------------------------------------------------------

class TestAddSheetXlsxEdit:

    def test_add_sheet_creates_new_sheet(self, lib):
        lib.load_workbook(XLSX_FILE)
        original = lib.list_sheet_names()
        lib.add_sheet("NewSheet")
        assert "NewSheet" in lib.list_sheet_names()

    def test_add_sheet_does_not_remove_existing_sheets(self, lib):
        lib.load_workbook(XLSX_FILE)
        original = lib.list_sheet_names()
        lib.add_sheet("Extra")
        updated = lib.list_sheet_names()
        for name in original:
            assert name in updated

    def test_add_sheet_switches_active_sheet(self, lib):
        lib.load_workbook(XLSX_FILE)
        lib.add_sheet("ActiveAfterAdd")
        rows = lib.get_rows()
        # Newly created sheet is empty — expect empty list
        assert rows == []

    def test_add_multiple_sheets(self, lib):
        lib.load_workbook(XLSX_FILE)
        lib.add_sheet("Alpha")
        lib.add_sheet("Beta")
        names = lib.list_sheet_names()
        assert "Alpha" in names
        assert "Beta" in names

    def test_add_sheet_increments_sheet_count(self, lib):
        lib.load_workbook(XLSX_FILE)
        before = len(lib.list_sheet_names())
        lib.add_sheet("OneMore")
        after = len(lib.list_sheet_names())
        assert after == before + 1


# ---------------------------------------------------------------------------
# XLSX – Streaming mode
# ---------------------------------------------------------------------------

class TestAddSheetXlsxStream:

    def test_add_sheet_raises_in_stream_mode(self, lib):
        lib.load_workbook(XLSX_FILE, read_only=True)
        with pytest.raises(LibraryException):
            lib.add_sheet("ShouldFail")


# ---------------------------------------------------------------------------
# XLS – Edit mode (lazy xls→xlsx conversion)
# ---------------------------------------------------------------------------

class TestAddSheetXlsEdit:

    def test_add_sheet_triggers_xls_conversion(self, lib):
        lib.load_workbook(XLS_FILE)
        lib.add_sheet("ConvertedSheet")
        assert "ConvertedSheet" in lib.list_sheet_names()

    def test_add_sheet_preserves_original_sheets_after_conversion(self, lib):
        lib.load_workbook(XLS_FILE)
        original = lib.list_sheet_names()
        lib.add_sheet("New")
        for name in original:
            assert name in lib.list_sheet_names()

    def test_add_sheet_new_sheet_is_empty(self, lib):
        lib.load_workbook(XLS_FILE)
        lib.add_sheet("EmptySheet")
        rows = lib.get_rows()
        assert rows == []

    def test_add_multiple_sheets_on_xls(self, lib):
        lib.load_workbook(XLS_FILE)
        lib.add_sheet("First")
        lib.add_sheet("Second")
        names = lib.list_sheet_names()
        assert "First" in names
        assert "Second" in names

    def test_add_sheet_increments_sheet_count(self, lib):
        lib.load_workbook(XLS_FILE)
        before = len(lib.list_sheet_names())
        lib.add_sheet("OneMore")
        after = len(lib.list_sheet_names())
        assert after == before + 1


# ---------------------------------------------------------------------------
# XLS – On-demand / streaming mode
# ---------------------------------------------------------------------------

class TestAddSheetXlsOnDemand:

    def test_add_sheet_raises_in_stream_mode(self, lib):
        lib.load_workbook(XLS_FILE, read_only=True)
        with pytest.raises(LibraryException):
            lib.add_sheet("ShouldFail")


# ---------------------------------------------------------------------------
# CSV – no sheet concept
# ---------------------------------------------------------------------------

class TestAddSheetCsv:

    def test_add_sheet_raises_for_csv_edit(self, lib):
        lib.load_workbook(CSV_FILE)
        with pytest.raises(OperationNotSupportedForFormat):
            lib.add_sheet("ShouldFail")

    def test_add_sheet_raises_for_csv_stream(self, lib):
        lib.load_workbook(CSV_FILE, read_only=True)
        with pytest.raises(LibraryException):
            lib.add_sheet("ShouldFail")
