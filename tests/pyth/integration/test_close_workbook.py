from pathlib import Path

import pytest

from rfexcel.exception.library_exceptions import WorkbookNotOpenException
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import BACKEND_NAMES, EDITABLE_FORMAT_LIST, open_backend


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_close_makes_loaded_workbook_inaccessible_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    lib.close()

    with pytest.raises(WorkbookNotOpenException):
        lib.get_rows()


@pytest.mark.parametrize("format_name", EDITABLE_FORMAT_LIST, ids=EDITABLE_FORMAT_LIST)
def test_close_makes_created_workbook_inaccessible_for_all_supported_formats(
    lib: RFExcelLibrary,
    format_name: str,
    tmp_path: Path,
) -> None:
    lib.create_workbook(str(tmp_path / f"new.{format_name}"))
    lib.close()

    with pytest.raises(WorkbookNotOpenException):
        lib.get_rows()


def test_close_without_open_workbook_does_not_raise(lib: RFExcelLibrary) -> None:
    lib.close()


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_close_is_idempotent_for_loaded_workbooks(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    lib.close()
    lib.close()


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_workbook_can_be_reloaded_after_close_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    expected_rows = lib.get_rows()
    lib.close()

    open_backend(lib, backend_name)
    assert lib.get_rows() == expected_rows


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_end_test_listener_closes_active_workbook_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    lib.end_test("some test", {})

    with pytest.raises(WorkbookNotOpenException):
        lib.get_rows()


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_close_cycle_can_be_repeated_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    for _ in range(2):
        open_backend(lib, backend_name)
        assert lib.get_rows()
        lib.close()

        with pytest.raises(WorkbookNotOpenException):
            lib.get_rows()
