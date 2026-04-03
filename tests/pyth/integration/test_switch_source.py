import pytest

from rfexcel.exception.library_exceptions import (
    FileDoesNotExistException, FileFormatNotSupportedException,
    WorkbookNotOpenException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import (BACKEND_NAMES, BACKENDS, CSV_EDIT,
                                  SHEET1_EXPECTED_ROW_COUNT, SHEET1_ROWS, XLS_EDIT, XLSX_EDIT)

CHAINED_TARGETS = [XLSX_EDIT, XLS_EDIT, CSV_EDIT]


@pytest.mark.parametrize("source_backend", BACKEND_NAMES, ids=BACKEND_NAMES)
@pytest.mark.parametrize("target_backend", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_switch_source_replaces_active_workbook_for_all_backend_pairs(
    lib: RFExcelLibrary,
    source_backend: str,
    target_backend: str,
) -> None:
    source_path, source_read_only = BACKENDS[source_backend]
    lib.load_workbook(source_path, read_only=source_read_only)

    target_path, target_read_only = BACKENDS[target_backend]
    lib.switch_source(target_path, read_only=target_read_only)

    assert lib.get_rows() == SHEET1_ROWS


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_switch_to_same_source_keeps_data_accessible_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    path, read_only = BACKENDS[backend_name]
    lib.load_workbook(path, read_only=read_only)
    lib.switch_source(path, read_only=read_only)

    assert lib.get_rows() == SHEET1_ROWS


@pytest.mark.parametrize("source_backend", BACKEND_NAMES, ids=BACKEND_NAMES)
@pytest.mark.parametrize("target_backend", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_switch_after_explicit_close_opens_new_source_for_all_backend_pairs(
    lib: RFExcelLibrary,
    source_backend: str,
    target_backend: str,
) -> None:
    source_path, source_read_only = BACKENDS[source_backend]
    lib.load_workbook(source_path, read_only=source_read_only)
    lib.close()

    target_path, target_read_only = BACKENDS[target_backend]
    lib.switch_source(target_path, read_only=target_read_only)

    assert len(lib.get_rows()) == SHEET1_EXPECTED_ROW_COUNT


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_switch_with_no_prior_workbook_opens_requested_backend(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    path, read_only = BACKENDS[backend_name]
    lib.switch_source(path, read_only=read_only)

    assert lib.get_rows() == SHEET1_ROWS


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_switch_to_nonexistent_file_raises_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    path, read_only = BACKENDS[backend_name]
    lib.load_workbook(path, read_only=read_only)

    with pytest.raises(FileDoesNotExistException):
        lib.switch_source("/nonexistent/missing.xlsx")


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_failed_switch_leaves_no_active_workbook_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    path, read_only = BACKENDS[backend_name]
    lib.load_workbook(path, read_only=read_only)

    with pytest.raises(FileDoesNotExistException):
        lib.switch_source("/nonexistent/missing.xlsx")

    with pytest.raises(WorkbookNotOpenException):
        lib.get_rows()


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_switch_to_unsupported_extension_raises_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    path, read_only = BACKENDS[backend_name]
    lib.load_workbook(path, read_only=read_only)

    with pytest.raises(FileFormatNotSupportedException):
        lib.switch_source("/some/file.ods")


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_switch_with_no_prior_workbook_and_bad_path_raises_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    _ = backend_name
    with pytest.raises(FileDoesNotExistException):
        lib.switch_source("/nonexistent/missing.csv")