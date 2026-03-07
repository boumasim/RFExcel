"""Integration tests for the Switch Source keyword.

Switch Source is a convenience keyword that closes the current workbook
and opens a new one in a single step. It accepts the same kwargs as
Load Workbook (e.g. read_only).
"""
import pytest

from rfexcel.exception.library_exceptions import (
    FileDoesNotExistException, FileFormatNotSupportedException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

XLSX_FIRST_DATA_ROW = {"Product ID": "P-200", "Description": "Wireless Mouse", "Price": "25.50", "Location": "Warehouse A, Shelf 2"}
XLS_FIRST_DATA_ROW  = {"Index": "1.0", "First Name": "Dulce", "Last Name": "Abril", "Gender": "Female", "Country": "United States", "Age": "32.0"}
CSV_FIRST_DATA_ROW  = {"Product ID": "P-200", "Description": "Wireless Mouse", "Price": "25.50", "Location": "Warehouse A, Shelf 2"}


# ─── switching from an open workbook ──────────────────────────────────────────

class TestSwitchSourceFromOpen:

    def test_switch_xlsx_to_csv_sets_new_active_workbook(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        old = lib._active_workbook
        lib.switch_source(CSV_FILE)
        assert lib._active_workbook is not None
        assert lib._active_workbook is not old

    def test_switch_xlsx_to_xls_data_is_from_new_file(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.switch_source(XLS_FILE)
        rows = lib.get_rows()
        assert len(rows) == 9
        assert rows[0] == XLS_FIRST_DATA_ROW

    def test_switch_xls_to_xlsx_data_is_from_new_file(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        lib.switch_source(XLSX_FILE)
        rows = lib.get_rows()
        assert len(rows) == 4
        assert rows[0] == XLSX_FIRST_DATA_ROW

    def test_switch_xlsx_to_csv_data_is_from_new_file(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.switch_source(CSV_FILE)
        rows = lib.get_rows()
        assert len(rows) == 4
        assert rows[0] == CSV_FIRST_DATA_ROW

    def test_switch_csv_to_xlsx_data_is_from_new_file(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        lib.switch_source(XLSX_FILE)
        rows = lib.get_rows()
        assert rows[0] == XLSX_FIRST_DATA_ROW

    def test_switch_to_same_file_reopens_it(self, lib: RFExcelLibrary):
        """Switching to the same file should give a fresh workbook instance."""
        lib.load_workbook(XLSX_FILE)
        old = lib._active_workbook
        lib.switch_source(XLSX_FILE)
        assert lib._active_workbook is not old
        assert lib.get_rows()[0] == XLSX_FIRST_DATA_ROW

    def test_switch_closes_previous_file_handle(self, lib: RFExcelLibrary):
        """The old stream resource file handle must be closed before the switch."""
        lib.load_workbook(CSV_FILE, read_only=True)
        old_resource = lib._active_workbook._resource  # type: ignore[union-attr]
        lib.switch_source(XLSX_FILE)
        assert old_resource._handle.closed  # type: ignore[union-attr]

    def test_switch_with_read_only_opens_in_stream_mode(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.switch_source(XLSX_FILE, read_only=True)
        # stream mode: rows still readable
        assert len(lib.get_rows()) == 4

    def test_switch_xlsx_stream_to_csv_edit(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.switch_source(CSV_FILE)
        assert lib.get_rows() == lib.get_rows()   # edit mode: repeatable


# ─── switching from no open workbook ──────────────────────────────────────────

class TestSwitchSourceFromClosed:

    def test_switch_with_no_prior_workbook_opens_file(self, lib: RFExcelLibrary):
        """Switch Source must work even when no workbook was previously open."""
        lib.switch_source(XLSX_FILE)
        assert lib._active_workbook is not None

    def test_switch_with_no_prior_workbook_data_correct(self, lib: RFExcelLibrary):
        lib.switch_source(XLSX_FILE)
        assert lib.get_rows()[0] == XLSX_FIRST_DATA_ROW

    def test_switch_after_explicit_close(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.close()
        lib.switch_source(CSV_FILE)
        assert lib.get_rows()[0] == CSV_FIRST_DATA_ROW


# ─── chained switches ─────────────────────────────────────────────────────────

class TestSwitchSourceChained:

    def test_three_successive_switches_last_one_wins(self, lib: RFExcelLibrary):
        lib.switch_source(XLSX_FILE)
        lib.switch_source(XLS_FILE)
        lib.switch_source(CSV_FILE)
        rows = lib.get_rows()
        assert len(rows) == 4
        assert rows[0] == CSV_FIRST_DATA_ROW

    def test_switch_back_and_forth_produces_correct_data(self, lib: RFExcelLibrary):
        lib.switch_source(XLSX_FILE)
        xlsx_rows = lib.get_rows()
        lib.switch_source(CSV_FILE)
        csv_rows = lib.get_rows()
        lib.switch_source(XLSX_FILE)
        assert lib.get_rows() == xlsx_rows
        assert csv_rows[0] == CSV_FIRST_DATA_ROW


# ─── negative ─────────────────────────────────────────────────────────────────

class TestSwitchSourceNegative:

    def test_switch_to_nonexistent_file_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(FileDoesNotExistException):
            lib.switch_source("/nonexistent/missing.xlsx")

    def test_active_workbook_is_none_after_failed_switch(self, lib: RFExcelLibrary):
        """If the new file doesn't exist the old one must already be closed."""
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(FileDoesNotExistException):
            lib.switch_source("/nonexistent/missing.xlsx")
        assert lib._active_workbook is None

    def test_switch_to_unsupported_extension_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(FileFormatNotSupportedException):
            lib.switch_source("/some/file.ods")

    def test_switch_with_no_prior_workbook_and_bad_path_raises(self, lib: RFExcelLibrary):
        with pytest.raises(FileDoesNotExistException):
            lib.switch_source("/nonexistent/missing.csv")
