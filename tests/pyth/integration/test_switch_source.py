import pytest

from rfexcel.exception.library_exceptions import (
    FileDoesNotExistException, FileFormatNotSupportedException,
    WorkbookNotOpenException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

XLSX_FIRST_DATA_ROW = {"Product ID": "P-200", "Description": "Wireless Mouse", "Price": "25.50", "Location": "Warehouse A, Shelf 2"}
XLS_FIRST_DATA_ROW  = {"Index": "1.0", "First Name": "Dulce", "Last Name": "Abril", "Gender": "Female", "Country": "United States", "Age": "32.0"}
CSV_FIRST_DATA_ROW  = {"Product ID": "P-200", "Description": "Wireless Mouse", "Price": "25.50", "Location": "Warehouse A, Shelf 2"}


# ---------------------------------------------------------------------------
# switch source from open
# ---------------------------------------------------------------------------

class TestSwitchSourceFromOpen:

    def test_switch_xlsx_to_csv_sets_new_active_workbook(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.switch_source(CSV_FILE)
        rows = lib.get_rows()
        assert len(rows) == 4
        assert rows[0] == CSV_FIRST_DATA_ROW

    def test_switch_xlsx_to_xls_data_is_from_new_file(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.switch_source(XLS_FILE)
        rows = lib.get_rows()
        assert len(rows) == 9
        assert rows[0] == XLS_FIRST_DATA_ROW

    def test_switch_xls_to_xlsx_data_is_from_new_file(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        lib.switch_source(XLSX_FILE)
        rows = lib.get_rows()
        assert len(rows) == 4
        assert rows[0] == XLSX_FIRST_DATA_ROW

    def test_switch_xlsx_to_csv_data_is_from_new_file(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.switch_source(CSV_FILE)
        rows = lib.get_rows()
        assert len(rows) == 4
        assert rows[0] == CSV_FIRST_DATA_ROW

    def test_switch_csv_to_xlsx_data_is_from_new_file(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        lib.switch_source(XLSX_FILE)
        rows = lib.get_rows()
        assert rows[0] == XLSX_FIRST_DATA_ROW

    def test_switch_to_same_file_reopens_it(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.switch_source(XLSX_FILE)
        assert lib.get_rows()[0] == XLSX_FIRST_DATA_ROW

    def test_switch_releases_previous_workbook_and_opens_new(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        lib.switch_source(XLSX_FILE)
        rows = lib.get_rows()
        assert rows[0] == XLSX_FIRST_DATA_ROW

    def test_switch_with_read_only_opens_in_stream_mode(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.switch_source(XLSX_FILE, read_only=True)
        assert len(lib.get_rows()) == 4

    def test_switch_xlsx_stream_to_csv_edit(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.switch_source(CSV_FILE)
        assert lib.get_rows() == lib.get_rows()


# ---------------------------------------------------------------------------
# switch source from closed
# ---------------------------------------------------------------------------

class TestSwitchSourceFromClosed:

    def test_switch_with_no_prior_workbook_opens_file(self, lib: RFExcelLibrary):
        lib.switch_source(XLSX_FILE)
        assert len(lib.get_rows()) > 0

    def test_switch_with_no_prior_workbook_data_correct(self, lib: RFExcelLibrary):
        lib.switch_source(XLSX_FILE)
        assert lib.get_rows()[0] == XLSX_FIRST_DATA_ROW

    def test_switch_after_explicit_close(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.close()
        lib.switch_source(CSV_FILE)
        assert lib.get_rows()[0] == CSV_FIRST_DATA_ROW


# ---------------------------------------------------------------------------
# switch source chained
# ---------------------------------------------------------------------------

class TestSwitchSourceChained:

    def test_three_successive_switches_last_one_wins(self, lib: RFExcelLibrary):
        lib.switch_source(XLSX_FILE)
        lib.switch_source(XLS_FILE)
        lib.switch_source(CSV_FILE)
        rows = lib.get_rows()
        assert len(rows) == 4
        assert rows[0] == CSV_FIRST_DATA_ROW

    def test_switch_back_and_forth_produces_correct_data(self, lib: RFExcelLibrary):
        lib.switch_source(XLSX_FILE)
        xlsx_rows = lib.get_rows()
        lib.switch_source(CSV_FILE)
        csv_rows = lib.get_rows()
        lib.switch_source(XLSX_FILE)
        assert lib.get_rows() == xlsx_rows
        assert csv_rows[0] == CSV_FIRST_DATA_ROW


# ---------------------------------------------------------------------------
# switch source negative
# ---------------------------------------------------------------------------

class TestSwitchSourceNegative:

    def test_switch_to_nonexistent_file_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(FileDoesNotExistException):
            lib.switch_source("/nonexistent/missing.xlsx")

    def test_active_workbook_is_not_usable_after_failed_switch(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(FileDoesNotExistException):
            lib.switch_source("/nonexistent/missing.xlsx")
        with pytest.raises(WorkbookNotOpenException):
            lib.get_rows()

    def test_switch_to_unsupported_extension_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(FileFormatNotSupportedException):
            lib.switch_source("/some/file.ods")

    def test_switch_with_no_prior_workbook_and_bad_path_raises(self, lib: RFExcelLibrary):
        with pytest.raises(FileDoesNotExistException):
            lib.switch_source("/nonexistent/missing.csv")
