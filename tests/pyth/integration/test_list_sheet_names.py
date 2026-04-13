import pytest

from rfexcel.exception.library_exceptions import (
	OperationNotSupportedForFormat,
	WorkbookNotOpenException,
)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import (
	BACKEND_NAMES,
	CSV_EDIT,
	CSV_STREAM,
	SHEET_LIST,
	XLSX_EDIT,
	open_backend,
)

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
    assert lib.list_sheet_names() == SHEET_LIST


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
