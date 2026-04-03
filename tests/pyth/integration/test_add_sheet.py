import pytest

from rfexcel.exception.library_exceptions import (
    HeadersNotDeterminedException, NullComponentException,
    OperationNotSupportedForFormat)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import (BACKEND_NAMES, CSV_EDIT, CSV_STREAM,
                                  XLS_EDIT, XLS_ON_DEMAND, XLSX_EDIT,
                                  XLSX_STREAM, open_backend)

EXPECTED_ADD_SHEET_EXCEPTION_BY_BACKEND: dict[str, type[Exception] | None] = {
    XLSX_EDIT: None,
    XLS_EDIT: None,
    XLSX_STREAM: NullComponentException,
    XLS_ON_DEMAND: NullComponentException,
    CSV_EDIT: OperationNotSupportedForFormat,
    CSV_STREAM: NullComponentException,
}


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_add_sheet_creates_new_sheet_or_raises_expected_exception(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    expected_exception = EXPECTED_ADD_SHEET_EXCEPTION_BY_BACKEND[backend_name]

    if expected_exception is not None:
        with pytest.raises(expected_exception):
            lib.add_sheet("NewSheet")
        return

    before = lib.list_sheet_names()
    lib.add_sheet("NewSheet")
    after = lib.list_sheet_names()
    assert "NewSheet" in after
    assert len(after) == len(before) + 1


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_add_sheet_preserves_existing_sheet_names_when_supported(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    expected_exception = EXPECTED_ADD_SHEET_EXCEPTION_BY_BACKEND[backend_name]

    if expected_exception is not None:
        with pytest.raises(expected_exception):
            lib.add_sheet("Extra")
        return

    original_sheet_names = lib.list_sheet_names()
    lib.add_sheet("Extra")
    updated_sheet_names = lib.list_sheet_names()
    for sheet_name in original_sheet_names:
        assert sheet_name in updated_sheet_names


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_add_sheet_switches_to_new_empty_sheet_when_supported(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    expected_exception = EXPECTED_ADD_SHEET_EXCEPTION_BY_BACKEND[backend_name]

    if expected_exception is not None:
        with pytest.raises(expected_exception):
            lib.add_sheet("ActiveAfterAdd")
        return

    lib.add_sheet("ActiveAfterAdd")
    with pytest.raises(HeadersNotDeterminedException):
        lib.get_rows()


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_add_multiple_sheets_behaves_consistently_per_backend(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    expected_exception = EXPECTED_ADD_SHEET_EXCEPTION_BY_BACKEND[backend_name]

    if expected_exception is not None:
        with pytest.raises(expected_exception):
            lib.add_sheet("Alpha")
        return

    lib.add_sheet("Alpha")
    lib.add_sheet("Beta")
    names = lib.list_sheet_names()
    assert "Alpha" in names
    assert "Beta" in names
