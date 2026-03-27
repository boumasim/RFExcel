import pytest

from rfexcel.exception.library_exceptions import (
    LibraryException, OperationNotSupportedForFormat,
    SheetDoesNotExistException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

XLSX_SHEET1_FIRST_ROW = {"Product ID": "P-200", "Description": "Wireless Mouse", "Price": "25.50", "Location": "Warehouse A, Shelf 2"}
XLSX_SHEET2_FIRST_ROW = {"Product ID": "P-300", "Description": "Wireless Mouse", "Price": "25.50", "Location": "Warehouse A, Shelf 2"}

XLS_SHEET1_FIRST_ROW = {"Index": "1.0", "First Name": "Dulce", "Last Name": "Abril", "Gender": "Female", "Country": "United States", "Age": "32.0"}
XLS_SHEET2_FIRST_ROW = {"Index": "1.0", "Date": "43023.0", "Id": "1562.0"}

XLS_SHEET1_HEADERS = ["Index", "First Name", "Last Name", "Gender", "Country", "Age"]
XLS_SHEET2_HEADERS = ["Index", "Date", "Id"]


# ---------------------------------------------------------------------------
# xlsx edit
# ---------------------------------------------------------------------------

class TestSwitchSheetXlsxEdit:

    def test_default_sheet_is_first_sheet(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows = lib.get_rows()
        assert rows[0] == XLSX_SHEET1_FIRST_ROW

    def test_switch_to_sheet2_changes_data(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.switch_sheet("Sheet2")
        rows = lib.get_rows()
        assert rows[0] == XLSX_SHEET2_FIRST_ROW

    def test_switch_to_sheet2_correct_row_count(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.switch_sheet("Sheet2")
        assert len(lib.get_rows()) == 4

    def test_switch_back_to_sheet1_restores_data(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.switch_sheet("Sheet2")
        lib.switch_sheet("List 1")
        rows = lib.get_rows()
        assert rows[0] == XLSX_SHEET1_FIRST_ROW

    def test_switch_does_not_affect_sheet_list(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.switch_sheet("Sheet2")
        assert lib.list_sheet_names() == ["List 1", "Sheet2", "Sheet3"]


# ---------------------------------------------------------------------------
# xlsx stream
# ---------------------------------------------------------------------------

class TestSwitchSheetXlsxStream:

    def test_default_sheet_is_first_sheet(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        rows = lib.get_rows()
        assert rows[0] == XLSX_SHEET1_FIRST_ROW

    def test_switch_to_sheet2_changes_data(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.switch_sheet("Sheet2")
        rows = lib.get_rows()
        assert rows[0] == XLSX_SHEET2_FIRST_ROW

    def test_switch_to_sheet2_correct_row_count(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.switch_sheet("Sheet2")
        assert len(lib.get_rows()) == 4

    def test_switch_resets_stream_position(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.get_rows()
        lib.switch_sheet("Sheet2")
        rows = lib.get_rows()
        assert rows[0] == XLSX_SHEET2_FIRST_ROW

    def test_switch_back_to_sheet1_resets_stream(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.switch_sheet("Sheet2")
        lib.switch_sheet("List 1")
        rows = lib.get_rows()
        assert rows[0] == XLSX_SHEET1_FIRST_ROW


# ---------------------------------------------------------------------------
# xls edit
# ---------------------------------------------------------------------------

class TestSwitchSheetXlsEdit:

    def test_default_sheet_is_first_sheet(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        rows = lib.get_rows()
        assert rows[0] == XLS_SHEET1_FIRST_ROW

    def test_switch_to_second_changes_data(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        lib.switch_sheet("Second")
        rows = lib.get_rows()
        assert rows[0] == XLS_SHEET2_FIRST_ROW

    def test_switch_to_second_correct_row_count(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        lib.switch_sheet("Second")
        assert len(lib.get_rows()) == 9

    def test_switch_to_second_correct_headers(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        lib.switch_sheet("Second")
        rows = lib.get_rows()
        assert list(rows[0].keys()) == XLS_SHEET2_HEADERS

    def test_switch_back_to_first_restores_data(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        lib.switch_sheet("Second")
        lib.switch_sheet("First")
        rows = lib.get_rows()
        assert rows[0] == XLS_SHEET1_FIRST_ROW

    def test_default_sheet_headers(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        rows = lib.get_rows()
        assert list(rows[0].keys()) == XLS_SHEET1_HEADERS


# ---------------------------------------------------------------------------
# xls on demand
# ---------------------------------------------------------------------------

class TestSwitchSheetXlsOnDemand:

    def test_default_sheet_is_first_sheet(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE, read_only=True)
        rows = lib.get_rows()
        assert rows[0] == XLS_SHEET1_FIRST_ROW

    def test_switch_to_second_changes_data(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE, read_only=True)
        lib.switch_sheet("Second")
        rows = lib.get_rows()
        assert rows[0] == XLS_SHEET2_FIRST_ROW

    def test_switch_to_second_correct_row_count(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE, read_only=True)
        lib.switch_sheet("Second")
        assert len(lib.get_rows()) == 9

    def test_switch_back_to_first_restores_data(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE, read_only=True)
        lib.switch_sheet("Second")
        lib.switch_sheet("First")
        rows = lib.get_rows()
        assert rows[0] == XLS_SHEET1_FIRST_ROW


# ---------------------------------------------------------------------------
# csv edit
# ---------------------------------------------------------------------------

class TestSwitchSheetCsv:

    def test_csv_edit_raises_operation_not_supported(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        with pytest.raises(OperationNotSupportedForFormat):
            lib.switch_sheet("anything")

    def test_csv_stream_raises_operation_not_supported(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        with pytest.raises(OperationNotSupportedForFormat):
            lib.switch_sheet("anything")

    def test_csv_exception_message_mentions_csv(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        with pytest.raises(OperationNotSupportedForFormat, match="(?i)csv"):
            lib.switch_sheet("anything")


# ---------------------------------------------------------------------------
# negative
# ---------------------------------------------------------------------------

class TestSwitchSheetNegative:

    def test_switch_to_nonexistent_sheet_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(SheetDoesNotExistException):
            lib.switch_sheet("DoesNotExist")

    def test_switch_to_nonexistent_sheet_xlsx_message(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(SheetDoesNotExistException, match="DoesNotExist"):
            lib.switch_sheet("DoesNotExist")

    def test_switch_to_nonexistent_sheet_xlsx_stream_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        with pytest.raises(SheetDoesNotExistException):
            lib.switch_sheet("DoesNotExist")

    def test_switch_to_nonexistent_sheet_xlsx_stream_message(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        with pytest.raises(SheetDoesNotExistException, match="DoesNotExist"):
            lib.switch_sheet("DoesNotExist")

    def test_switch_to_nonexistent_sheet_xls_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        with pytest.raises(SheetDoesNotExistException):
            lib.switch_sheet("DoesNotExist")

    def test_switch_to_nonexistent_sheet_xls_message(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        with pytest.raises(SheetDoesNotExistException, match="DoesNotExist"):
            lib.switch_sheet("DoesNotExist")

    def test_switch_to_nonexistent_sheet_xls_on_demand_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE, read_only=True)
        with pytest.raises(SheetDoesNotExistException):
            lib.switch_sheet("DoesNotExist")

    def test_switch_to_nonexistent_sheet_xls_on_demand_message(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE, read_only=True)
        with pytest.raises(SheetDoesNotExistException, match="DoesNotExist"):
            lib.switch_sheet("DoesNotExist")

    def test_switch_to_nonexistent_sheet_csv_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        with pytest.raises(OperationNotSupportedForFormat):
            lib.switch_sheet("DoesNotExist")            lib.switch_sheet("DoesNotExist")