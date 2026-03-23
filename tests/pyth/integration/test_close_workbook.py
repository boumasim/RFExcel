from pathlib import Path
import pytest

from rfexcel.exception.library_exceptions import WorkbookNotOpenException
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

# ---------------------------------------------------------------------------
# Positive
# ---------------------------------------------------------------------------

class TestCloseWorkbookPositive:

    def test_close_after_load_xlsx_makes_workbook_inaccessible(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.close()
        with pytest.raises(WorkbookNotOpenException):
            lib.get_rows()

    def test_close_after_load_xlsx_stream_makes_workbook_inaccessible(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.close()
        with pytest.raises(WorkbookNotOpenException):
            lib.get_rows()

    def test_close_after_load_xls_makes_workbook_inaccessible(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        lib.close()
        with pytest.raises(WorkbookNotOpenException):
            lib.get_rows()

    def test_close_after_load_csv_makes_workbook_inaccessible(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        lib.close()
        with pytest.raises(WorkbookNotOpenException):
            lib.get_rows()

    def test_close_after_create_xlsx_makes_workbook_inaccessible(self, lib: RFExcelLibrary, tmp_path: Path):
        lib.create_workbook(str(tmp_path / "new.xlsx"))
        lib.close()
        with pytest.raises(WorkbookNotOpenException):
            lib.get_rows()

    def test_close_after_create_csv_makes_workbook_inaccessible(self, lib: RFExcelLibrary, tmp_path: Path):
        lib.create_workbook(str(tmp_path / "new.csv"))
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

    def test_close_csv_stream_makes_workbook_inaccessible(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        lib.close()
        with pytest.raises(WorkbookNotOpenException):
            lib.get_rows()

    def test_close_then_reload_then_close_again(self, lib: RFExcelLibrary):
        for _ in range(2):
            lib.load_workbook(XLSX_FILE)
            assert len(lib.get_rows()) > 0
            lib.close()
            with pytest.raises(WorkbookNotOpenException):
                lib.get_rows()
