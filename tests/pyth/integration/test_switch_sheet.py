"""Integration tests for the Switch Sheet keyword.

Expected sheet data derived from test resources:

data.xlsx:
  'List 1' (default, sheet index 0):
    Headers : Product ID | Description | Price | Location
    Row 2   : P-200 | Wireless Mouse | 25.50 | Warehouse A, Shelf 2
  'Sheet2' (sheet index 1):
    Headers : Product ID | Description | Price | Location
    Row 2   : P-300 | Wireless Mouse | 25.50 | Warehouse A, Shelf 2

example.xls:
  'First' (default, sheet index 0):
    Headers : Index | First Name | Last Name | Gender | Country | Age | '' | ''
    Row 2   : 1.0 | Dulce | Abril | Female | United States | 32.0 | '' | ''
  'Second' (sheet index 1):
    Headers : Index | Date | Id
    Row 2   : 1.0 | 43023.0 | 1562.0

data.csv: raises OperationNotSupportedForFormat (no sheet concept)
"""
import pytest

from rfexcel.exception.library_exceptions import OperationNotSupportedForFormat
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

# ─── expected data ─────────────────────────────────────────────────────────────

XLSX_SHEET1_FIRST_ROW = {"Product ID": "P-200", "Description": "Wireless Mouse", "Price": "25.50", "Location": "Warehouse A, Shelf 2"}
XLSX_SHEET2_FIRST_ROW = {"Product ID": "P-300", "Description": "Wireless Mouse", "Price": "25.50", "Location": "Warehouse A, Shelf 2"}

XLS_SHEET1_FIRST_ROW = {"Index": "1.0", "First Name": "Dulce", "Last Name": "Abril", "Gender": "Female", "Country": "United States", "Age": "32.0", "": ""}
XLS_SHEET2_FIRST_ROW = {"Index": "1.0", "Date": "43023.0", "Id": "1562.0"}

XLS_SHEET1_HEADERS = ["Index", "First Name", "Last Name", "Gender", "Country", "Age", ""]
XLS_SHEET2_HEADERS = ["Index", "Date", "Id"]


# ─── xlsx edit mode ─────────────────────────────────────────────────────────────

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
        assert lib.list_sheet_names() == ["List 1", "Sheet2"]


# ─── xlsx streaming mode ──────────────────────────────────────────────────────

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
        """Switching sheet resets the row generator so row 1 is read again."""
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.get_rows()  # exhaust the stream on List 1
        lib.switch_sheet("Sheet2")
        rows = lib.get_rows()  # must read from row 1 of Sheet2
        assert rows[0] == XLSX_SHEET2_FIRST_ROW

    def test_switch_back_to_sheet1_resets_stream(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.switch_sheet("Sheet2")
        lib.switch_sheet("List 1")
        rows = lib.get_rows()
        assert rows[0] == XLSX_SHEET1_FIRST_ROW


# ─── xls standard (edit) mode ─────────────────────────────────────────────────

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


# ─── xls on-demand (streaming) mode ──────────────────────────────────────────

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


# ─── csv (unsupported) ────────────────────────────────────────────────────────

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
