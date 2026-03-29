import pytest

from rfexcel.exception.library_exceptions import (
    OperationNotSupportedForFormat, WorkbookNotOpenException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

XLSX_SHEET_NAMES = ["List 1", "Sheet2", "Sheet3", "Sheet4"]
XLS_SHEET_NAMES  = ["First", "Second"]


# ---------------------------------------------------------------------------
# XLSX and XLS – all modes
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    ("path", "read_only", "expected"),
    [
        (XLSX_FILE, False, XLSX_SHEET_NAMES),
        (XLSX_FILE, True,  XLSX_SHEET_NAMES),
        (XLS_FILE,  False, XLS_SHEET_NAMES),
        (XLS_FILE,  True,  XLS_SHEET_NAMES),
    ],
    ids=["xlsx_edit", "xlsx_stream", "xls_edit", "xls_on_demand"],
)
def test_returns_correct_sheet_names(
    lib: RFExcelLibrary, path: str, read_only: bool, expected: list[str]
):
    lib.load_workbook(path, read_only=read_only)
    assert lib.list_sheet_names() == expected


@pytest.mark.parametrize(
    ("path", "read_only"),
    [
        (XLSX_FILE, False),
        (XLSX_FILE, True),
        (XLS_FILE,  False),
        (XLS_FILE,  True),
    ],
    ids=["xlsx_edit", "xlsx_stream", "xls_edit", "xls_on_demand"],
)
def test_returns_list_type(lib: RFExcelLibrary, path: str, read_only: bool):
    lib.load_workbook(path, read_only=read_only)
    assert isinstance(lib.list_sheet_names(), list)


# ---------------------------------------------------------------------------
# CSV – raises for both modes
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "read_only",
    [False, True],
    ids=["csv_edit", "csv_stream"],
)
def test_csv_raises_operation_not_supported(lib: RFExcelLibrary, read_only: bool):
    lib.load_workbook(CSV_FILE, read_only=read_only)
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
