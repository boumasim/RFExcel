import pytest

from rfexcel.exception.library_exceptions import (
    FileDoesNotExistException, FileFormatNotSupportedException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

# ---------------------------------------------------------------------------
# Shared parametrize data
# ---------------------------------------------------------------------------

_ALL_FILES_AND_MODES = [
    (XLSX_FILE, False),
    (XLSX_FILE, True),
    (XLS_FILE,  False),
    (XLS_FILE,  True),
    (CSV_FILE,  False),
    (CSV_FILE,  True),
]
_ALL_FILES_AND_MODES_IDS = [
    "xlsx_edit", "xlsx_stream",
    "xls_edit",  "xls_on_demand",
    "csv_edit",  "csv_stream",
]


# ---------------------------------------------------------------------------
# load workbook positive
# ---------------------------------------------------------------------------

class TestLoadWorkbookPositive:

    @pytest.mark.parametrize(("path", "read_only"), _ALL_FILES_AND_MODES, ids=_ALL_FILES_AND_MODES_IDS)
    def test_sets_active_workbook(self, lib: RFExcelLibrary, path: str, read_only: bool):
        lib.load_workbook(path, read_only=read_only)
        assert lib._active_workbook is not None

    @pytest.mark.parametrize(
        ("path", "read_only", "expected_rows"),
        [
            (XLSX_FILE, False, 4),
            (XLSX_FILE, True,  4),
            (XLS_FILE,  False, 9),
            (XLS_FILE,  True,  9),
            (CSV_FILE,  False, 4),
            (CSV_FILE,  True,  4),
        ],
        ids=_ALL_FILES_AND_MODES_IDS,
    )
    def test_is_immediately_readable(
        self, lib: RFExcelLibrary, path: str, read_only: bool, expected_rows: int
    ):
        lib.load_workbook(path, read_only=read_only)
        assert len(lib.get_rows()) == expected_rows

    def test_load_xlsx_edit_first_row_content(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows()
        assert rows[0]["Product ID"] == "P-200"


# ---------------------------------------------------------------------------
# negative / edge
# ---------------------------------------------------------------------------

class TestLoadWorkbookNegative:

    @pytest.mark.parametrize("path", [
        "/nonexistent/path/missing.xlsx",
        "/nonexistent/path/missing.csv",
        "/nonexistent/path/missing.xls",
    ], ids=["xlsx", "csv", "xls"])
    def test_non_existent_file_raises(self, lib: RFExcelLibrary, path: str):
        with pytest.raises(FileDoesNotExistException):
            lib.load_workbook(path)

    @pytest.mark.parametrize("path", [
        "/some/path/file.txt",
        "/some/path/file.ods",
    ], ids=["txt", "ods"])
    def test_unsupported_extension_raises(self, lib: RFExcelLibrary, path: str):
        with pytest.raises(FileFormatNotSupportedException):
            lib.load_workbook(path)

    def test_active_workbook_is_none_after_failed_load(self, lib: RFExcelLibrary):
        with pytest.raises(FileDoesNotExistException):
            lib.load_workbook("/nonexistent/path/missing.xlsx")
        assert lib._active_workbook is None

class TestLoadWorkbookEdge:

    def test_loading_second_file_replaces_first(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        first_wb = lib._active_workbook
        lib.load_workbook(CSV_FILE)
        assert lib._active_workbook is not first_wb

    def test_loading_after_close_works(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.close()
        lib.load_workbook(XLSX_FILE)
        assert lib._active_workbook is not None
        assert len(lib.get_rows()) == 4

    def test_xlsx_edit_and_stream_produce_identical_rows(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        edit_rows = lib.get_rows()
        lib.close()
        lib.load_workbook(XLSX_FILE, read_only=True)
        stream_rows = lib.get_rows()
        assert edit_rows == stream_rows

    def test_csv_edit_and_stream_produce_identical_rows(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        edit_rows = lib.get_rows()
        lib.close()
        lib.load_workbook(CSV_FILE, read_only=True)
        stream_rows = lib.get_rows()
        assert edit_rows == stream_rows

    def test_xls_edit_and_on_demand_produce_identical_rows(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        edit_rows = lib.get_rows()
        lib.close()
        lib.load_workbook(XLS_FILE, read_only=True)
        on_demand_rows = lib.get_rows()
        assert edit_rows == on_demand_rows
