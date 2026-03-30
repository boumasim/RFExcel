import pytest

from rfexcel.exception.library_exceptions import (
    LibraryException, NullComponentException, OperationNotSupportedForFormat)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import (BACKEND_NAMES, CSV_EDIT, CSV_STREAM,
                                  XLS_EDIT, XLS_ON_DEMAND, XLSX_EDIT,
                                  XLSX_STREAM, open_backend)

EXPECTED_DELETE_SHEET_EXCEPTION_BY_BACKEND: dict[str, type[Exception] | None] = {
    XLSX_EDIT: None,
    XLS_EDIT: None,
    CSV_EDIT: OperationNotSupportedForFormat,
    XLSX_STREAM: NullComponentException,
    XLS_ON_DEMAND: NullComponentException,
    CSV_STREAM: NullComponentException,
}


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_delete_sheet_removes_sheet_when_supported_or_raises_expected_exception(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    expected_exception = EXPECTED_DELETE_SHEET_EXCEPTION_BY_BACKEND[backend_name]

    if expected_exception is not None:
        with pytest.raises(expected_exception):
            lib.delete_sheet("ToDelete")
        return

    lib.add_sheet("ToDelete")
    lib.delete_sheet("ToDelete")
    assert "ToDelete" not in lib.list_sheet_names()


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_delete_sheet_decrements_sheet_count_when_supported(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    expected_exception = EXPECTED_DELETE_SHEET_EXCEPTION_BY_BACKEND[backend_name]

    if expected_exception is not None:
        with pytest.raises(expected_exception):
            lib.delete_sheet("Temp")
        return

    lib.add_sheet("Temp")
    before = len(lib.list_sheet_names())
    lib.delete_sheet("Temp")
    assert len(lib.list_sheet_names()) == before - 1


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_delete_sheet_preserves_other_sheets_when_supported(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    expected_exception = EXPECTED_DELETE_SHEET_EXCEPTION_BY_BACKEND[backend_name]

    if expected_exception is not None:
        with pytest.raises(expected_exception):
            lib.delete_sheet("Remove")
        return

    lib.add_sheet("Keep")
    lib.add_sheet("Remove")
    lib.delete_sheet("Remove")
    assert "Keep" in lib.list_sheet_names()


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_delete_sheet_resets_active_to_first_sheet_when_supported(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    expected_exception = EXPECTED_DELETE_SHEET_EXCEPTION_BY_BACKEND[backend_name]

    if expected_exception is not None:
        with pytest.raises(expected_exception):
            lib.delete_sheet("Victim")
        return

    first_sheet = lib.list_sheet_names()[0]
    lib.add_sheet("Victim")
    lib.delete_sheet("Victim")
    assert lib.list_sheet_names()[0] == first_sheet
    rows = lib.get_rows()
    assert isinstance(rows, list)
    assert len(rows) > 0


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_delete_nonexistent_sheet_raises_expected_exception(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    expected_exception = EXPECTED_DELETE_SHEET_EXCEPTION_BY_BACKEND[backend_name]

    if expected_exception is not None:
        with pytest.raises(expected_exception):
            lib.delete_sheet("DoesNotExist")
        return

    with pytest.raises(LibraryException):
        lib.delete_sheet("DoesNotExist")
