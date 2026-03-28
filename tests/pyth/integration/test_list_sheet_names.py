import pytest

from rfexcel.exception.library_exceptions import (
    OperationNotSupportedForFormat, WorkbookNotOpenException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

XLSX_SHEET_NAMES = ["List 1", "Sheet2", "Sheet3", "Sheet4"]
XLS_SHEET_NAMES  = ["First", "Second"]


# ---------------------------------------------------------------------------
# XLSX edit
# ---------------------------------------------------------------------------

class TestListSheetNamesXlsxEdit:

    def test_returns_correct_sheet_names(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert lib.list_sheet_names() == XLSX_SHEET_NAMES

    def test_returns_list_type(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        assert isinstance(lib.list_sheet_names(), list)


# ---------------------------------------------------------------------------
# XLSX stream
# ---------------------------------------------------------------------------

class TestListSheetNamesXlsxStream:

    def test_returns_correct_sheet_names(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        assert lib.list_sheet_names() == XLSX_SHEET_NAMES

    def test_returns_list_type(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        assert isinstance(lib.list_sheet_names(), list)


# ---------------------------------------------------------------------------
# XLS edit
# ---------------------------------------------------------------------------

class TestListSheetNamesXlsEdit:

    def test_returns_correct_sheet_names(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        assert lib.list_sheet_names() == XLS_SHEET_NAMES

    def test_returns_list_type(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        assert isinstance(lib.list_sheet_names(), list)


# ---------------------------------------------------------------------------
# XLS stream / on demand
# ---------------------------------------------------------------------------

class TestListSheetNamesXlsOnDemand:

    def test_returns_correct_sheet_names(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE, read_only=True)
        assert lib.list_sheet_names() == XLS_SHEET_NAMES

    def test_returns_list_type(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE, read_only=True)
        assert isinstance(lib.list_sheet_names(), list)


# ---------------------------------------------------------------------------
# CSV – edit
# ---------------------------------------------------------------------------

class TestListSheetNamesCsvEdit:

    def test_raises_operation_not_supported(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        with pytest.raises(OperationNotSupportedForFormat):
            lib.list_sheet_names()


# ---------------------------------------------------------------------------
# CSV – stream
# ---------------------------------------------------------------------------

class TestListSheetNamesCsvStream:

    def test_raises_operation_not_supported(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        with pytest.raises(OperationNotSupportedForFormat):
            lib.list_sheet_names()


# ---------------------------------------------------------------------------
# No active workbook
# ---------------------------------------------------------------------------

class TestListSheetNamesNoWorkbook:

    def test_raises_when_no_workbook_open(self, lib: RFExcelLibrary):
        with pytest.raises(WorkbookNotOpenException):
            lib.list_sheet_names()

    def test_raises_after_close(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.close()
        with pytest.raises(WorkbookNotOpenException):
            lib.list_sheet_names()
