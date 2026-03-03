"""Integration tests for the List Sheet Names keyword.

Expected sheet names derived from test resources:
  data.xlsx  : ['List 1', 'Sheet2']
  example.xls: ['First', 'Second']
  data.csv   : raises OperationNotSupportedForFormat (no sheet concept)

Covers:
  - xlsx edit and streaming mode.
  - xls standard and on-demand mode.
  - CSV edit and streaming mode → raises OperationNotSupportedForFormat.
  - No active workbook → returns empty list.
"""
import pytest

from rfexcel.exception.library_exceptions import LibraryException, OperationNotSupportedForFormat
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

XLSX_SHEET_NAMES = ["List 1", "Sheet2"]
XLS_SHEET_NAMES  = ["First", "Second"]


# ─── xlsx ─────────────────────────────────────────────────────────────────────

class TestListSheetNamesXlsx:

    def test_xlsx_edit_returns_correct_sheet_names(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert lib.list_sheet_names() == XLSX_SHEET_NAMES

    def test_xlsx_edit_returns_list_type(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert isinstance(lib.list_sheet_names(), list)

    def test_xlsx_stream_returns_correct_sheet_names(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        assert lib.list_sheet_names() == XLSX_SHEET_NAMES

    def test_xlsx_stream_returns_list_type(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        assert isinstance(lib.list_sheet_names(), list)


# ─── xls ──────────────────────────────────────────────────────────────────────

class TestListSheetNamesXls:

    def test_xls_edit_returns_correct_sheet_names(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        assert lib.list_sheet_names() == XLS_SHEET_NAMES

    def test_xls_edit_returns_list_type(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        assert isinstance(lib.list_sheet_names(), list)

    def test_xls_on_demand_returns_correct_sheet_names(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE, read_only=True)
        assert lib.list_sheet_names() == XLS_SHEET_NAMES

    def test_xls_on_demand_returns_list_type(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE, read_only=True)
        assert isinstance(lib.list_sheet_names(), list)


# ─── csv (unsupported) ────────────────────────────────────────────────────────

class TestListSheetNamesCsv:

    def test_csv_edit_raises_operation_not_supported(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        with pytest.raises(LibraryException):
            lib.list_sheet_names()

    def test_csv_stream_raises_operation_not_supported(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        with pytest.raises(LibraryException):
            lib.list_sheet_names()


# ─── no active workbook ───────────────────────────────────────────────────────

class TestListSheetNamesNoWorkbook:

    def test_returns_empty_list_when_no_workbook_open(self, lib: RFExcelLibrary):
        assert lib.list_sheet_names() == []

    def test_returns_empty_list_after_close(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.close()
        assert lib.list_sheet_names() == []
