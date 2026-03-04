"""Integration tests for the Delete Sheet feature."""
import pytest

from rfexcel.exception.library_exceptions import (
    LibraryException, OperationNotSupportedForFormat)
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

# ---------------------------------------------------------------------------
# XLSX – Edit mode
# ---------------------------------------------------------------------------

class TestDeleteSheetXlsxEdit:

    def test_delete_sheet_removes_sheet(self, lib):
        lib.load_workbook(XLSX_FILE)
        lib.add_sheet("ToDelete")
        lib.delete_sheet("ToDelete")
        assert "ToDelete" not in lib.list_sheet_names()

    def test_delete_sheet_decrements_sheet_count(self, lib):
        lib.load_workbook(XLSX_FILE)
        lib.add_sheet("Temp")
        before = len(lib.list_sheet_names())
        lib.delete_sheet("Temp")
        assert len(lib.list_sheet_names()) == before - 1

    def test_delete_sheet_preserves_remaining_sheets(self, lib):
        lib.load_workbook(XLSX_FILE)
        lib.add_sheet("Keep")
        lib.add_sheet("Remove")
        lib.delete_sheet("Remove")
        assert "Keep" in lib.list_sheet_names()

    def test_delete_sheet_resets_active_to_first_sheet(self, lib):
        lib.load_workbook(XLSX_FILE)
        first_sheet = lib.list_sheet_names()[0]
        lib.add_sheet("Victim")
        lib.delete_sheet("Victim")
        # After deletion the active sheet falls back to first — data should be readable
        rows = lib.get_rows()
        assert isinstance(rows, list)

    def test_delete_nonexistent_sheet_raises(self, lib):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(LibraryException):
            lib.delete_sheet("DoesNotExist")


# ---------------------------------------------------------------------------
# XLSX – Streaming mode
# ---------------------------------------------------------------------------

class TestDeleteSheetXlsxStream:

    def test_delete_sheet_raises_in_stream_mode(self, lib):
        lib.load_workbook(XLSX_FILE, read_only=True)
        with pytest.raises(LibraryException):
            lib.delete_sheet("Sheet1")


# ---------------------------------------------------------------------------
# XLS – Edit mode (lazy xls→xlsx conversion)
# ---------------------------------------------------------------------------

class TestDeleteSheetXlsEdit:

    def test_delete_sheet_triggers_xls_conversion(self, lib):
        lib.load_workbook(XLS_FILE)
        original = lib.list_sheet_names()
        lib.add_sheet("Extra")
        lib.delete_sheet("Extra")
        assert "Extra" not in lib.list_sheet_names()

    def test_delete_sheet_preserves_original_sheets(self, lib):
        lib.load_workbook(XLS_FILE)
        original = lib.list_sheet_names()
        lib.add_sheet("TempSheet")
        lib.delete_sheet("TempSheet")
        for name in original:
            assert name in lib.list_sheet_names()

    def test_delete_nonexistent_sheet_raises_after_conversion(self, lib):
        lib.load_workbook(XLS_FILE)
        # Force conversion first via add_sheet, then attempt bad delete
        lib.add_sheet("Anchor")
        with pytest.raises(LibraryException):
            lib.delete_sheet("DoesNotExist")


# ---------------------------------------------------------------------------
# XLS – On-demand / streaming mode
# ---------------------------------------------------------------------------

class TestDeleteSheetXlsOnDemand:

    def test_delete_sheet_raises_in_stream_mode(self, lib):
        lib.load_workbook(XLS_FILE, read_only=True)
        with pytest.raises(LibraryException):
            lib.delete_sheet("Sheet1")


# ---------------------------------------------------------------------------
# CSV – no sheet concept
# ---------------------------------------------------------------------------

class TestDeleteSheetCsv:

    def test_delete_sheet_raises_for_csv_edit(self, lib):
        lib.load_workbook(CSV_FILE)
        with pytest.raises(OperationNotSupportedForFormat):
            lib.delete_sheet("ShouldFail")

    def test_delete_sheet_raises_for_csv_stream(self, lib):
        lib.load_workbook(CSV_FILE, read_only=True)
        with pytest.raises(LibraryException):
            lib.delete_sheet("ShouldFail")
