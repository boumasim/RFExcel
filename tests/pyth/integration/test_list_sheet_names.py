import pytest

from rfexcel.exception.library_exceptions import (
    OperationNotSupportedForFormat, WorkbookNotOpenException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import (BACKEND_NAMES, CSV_EDIT, CSV_STREAM,
                                  XLS_EDIT, XLS_ON_DEMAND, XLSX_EDIT,
                                  XLSX_STREAM, open_backend)

EXPECTED_SHEET_NAMES_BY_BACKEND: dict[str, list[str]] = {
    XLSX_EDIT: ["List 1", "Sheet2", "Sheet3", "Sheet4"],
    XLSX_STREAM: ["List 1", "Sheet2", "Sheet3", "Sheet4"],
    XLS_EDIT: ["First", "Second"],
    XLS_ON_DEMAND: ["First", "Second"],
}

CSV_BACKENDS = [CSV_EDIT, CSV_STREAM]


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_returns_correct_sheet_names(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    if backend_name in CSV_BACKENDS:
        with pytest.raises(OperationNotSupportedForFormat):
            lib.list_sheet_names()
        return
    assert lib.list_sheet_names() == EXPECTED_SHEET_NAMES_BY_BACKEND[backend_name]


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_returns_list_type_for_supported_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    if backend_name in CSV_BACKENDS:
        with pytest.raises(OperationNotSupportedForFormat):
            lib.list_sheet_names()
        return
    assert isinstance(lib.list_sheet_names(), list)


def test_raises_when_no_workbook_open(lib: RFExcelLibrary) -> None:
    with pytest.raises(WorkbookNotOpenException):
        lib.list_sheet_names()


def test_raises_after_close(lib: RFExcelLibrary) -> None:
    open_backend(lib, XLSX_EDIT)
    lib.close()
    with pytest.raises(WorkbookNotOpenException):
        lib.list_sheet_names()
