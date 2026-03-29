from pathlib import Path

import pytest

from rfexcel.exception.library_exceptions import WorkbookNotOpenException
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

# ---------------------------------------------------------------------------
# Positive
# ---------------------------------------------------------------------------

class TestCloseWorkbookPositive:

    @pytest.mark.parametrize(
        ("path", "read_only"),
        [
            (XLSX_FILE, False),
            (XLSX_FILE, True),
            (XLS_FILE,  False),
            (CSV_FILE,  False),
            (CSV_FILE,  True),
        ],
        ids=["xlsx_edit", "xlsx_stream", "xls_edit", "csv_edit", "csv_stream"],
    )
    def test_close_after_load_makes_workbook_inaccessible(
        self, lib: RFExcelLibrary, path: str, read_only: bool
    ):
        lib.load_workbook(path, read_only=read_only)
        lib.close()
        with pytest.raises(WorkbookNotOpenException):
            lib.get_rows()

    @pytest.mark.parametrize(
        "filename",
        ["new.xlsx", "new.csv"],
        ids=["xlsx", "csv"],
    )
    def test_close_after_create_makes_workbook_inaccessible(
        self, lib: RFExcelLibrary, tmp_path: Path, filename: str
    ):
        lib.create_workbook(str(tmp_path / filename))
        lib.close()
        with pytest.raises(WorkbookNotOpenException):
            lib.get_rows()


# ---------------------------------------------------------------------------
# Negative / edge
# ---------------------------------------------------------------------------

class TestCloseWorkbookEdge:

    def test_close_without_open_does_not_raise(self, lib: RFExcelLibrary):
        lib.close()

    def test_close_twice_does_not_raise(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.close()
        lib.close()

    def test_get_rows_after_close_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.close()
        with pytest.raises(WorkbookNotOpenException):
            lib.get_rows()

    def test_reload_after_close_works(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.close()
        lib.load_workbook(XLSX_FILE)
        assert len(lib.get_rows()) == 4

    def test_listener_closes_workbook_automatically(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.end_test("some test", {})
        with pytest.raises(WorkbookNotOpenException):
            lib.get_rows()

    def test_close_then_reload_then_close_again(self, lib: RFExcelLibrary):
        for _ in range(2):
            lib.load_workbook(XLSX_FILE)
            assert len(lib.get_rows()) > 0
            lib.close()
            with pytest.raises(WorkbookNotOpenException):
                lib.get_rows()
