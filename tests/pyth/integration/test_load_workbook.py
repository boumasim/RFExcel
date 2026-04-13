import pytest

from rfexcel.exception.library_exceptions import (
	FileDoesNotExistException,
	FileFormatNotSupportedException,
	WorkbookNotOpenException,
)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import (
	BACKEND_NAMES,
	BACKENDS,
	CSV_EDIT,
	CSV_FORMAT,
	CSV_STREAM,
	SHEET1_EXPECTED_ROW_COUNT,
	XLS_EDIT,
	XLS_FORMAT,
	XLS_ON_DEMAND,
	XLSX_EDIT,
	XLSX_FORMAT,
	XLSX_STREAM,
)

EDIT_AND_READONLY_PAIRS: list[tuple[str, str]] = [
    (XLSX_EDIT, XLSX_STREAM),
    (CSV_EDIT, CSV_STREAM),
    (XLS_EDIT, XLS_ON_DEMAND),
]

@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_load_workbook_is_immediately_readable_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    path, read_only = BACKENDS[backend_name]
    lib.load_workbook(path, read_only=read_only)
    assert len(lib.get_rows()) == SHEET1_EXPECTED_ROW_COUNT


@pytest.mark.parametrize("path", [
    f"/nonexistent/path/missing.{XLSX_FORMAT}",
    f"/nonexistent/path/missing.{CSV_FORMAT}",
    f"/nonexistent/path/missing.{XLS_FORMAT}",
], ids=[XLSX_FORMAT, CSV_FORMAT, XLS_FORMAT])
def test_load_workbook_nonexistent_file_raises(
    lib: RFExcelLibrary,
    path: str,
) -> None:
    with pytest.raises(FileDoesNotExistException):
        lib.load_workbook(path)


@pytest.mark.parametrize("path", [
    "/some/path/file.txt",
    "/some/path/file.ods",
], ids=["txt", "ods"])
def test_load_workbook_unsupported_extension_raises(
    lib: RFExcelLibrary,
    path: str,
) -> None:
    with pytest.raises(FileFormatNotSupportedException):
        lib.load_workbook(path)


def test_active_workbook_is_none_after_failed_load(lib: RFExcelLibrary) -> None:
    with pytest.raises(FileDoesNotExistException):
        lib.load_workbook("/nonexistent/path/missing.xlsx")

    assert lib.active_workbook is None

    with pytest.raises(WorkbookNotOpenException):
        lib.get_rows()


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_loading_after_close_works_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    path, read_only = BACKENDS[backend_name]
    lib.load_workbook(path, read_only=read_only)
    lib.close()
    lib.load_workbook(path, read_only=read_only)

    assert len(lib.get_rows()) == SHEET1_EXPECTED_ROW_COUNT


@pytest.mark.parametrize(
    ("edit_backend", "readonly_backend"),
    EDIT_AND_READONLY_PAIRS,
    ids=[XLSX_FORMAT, CSV_FORMAT, XLS_FORMAT],
)
def test_edit_and_readonly_load_modes_produce_identical_rows(
    lib: RFExcelLibrary,
    edit_backend: str,
    readonly_backend: str,
) -> None:
    edit_path, edit_read_only = BACKENDS[edit_backend]
    lib.load_workbook(edit_path, read_only=edit_read_only)
    edit_rows = lib.get_rows()
    lib.close()

    readonly_path, readonly_read_only = BACKENDS[readonly_backend]
    lib.load_workbook(readonly_path, read_only=readonly_read_only)
    readonly_rows = lib.get_rows()

    assert edit_rows == readonly_rows
