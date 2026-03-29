from typing import Any, cast

import pytest

from rfexcel.exception.library_exceptions import (
    OperationNotSupportedForFormat, SheetDoesNotExistException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

XLSX_SHEET1_FIRST_ROW = {"Product ID": "P-200", "Description": "Wireless Mouse", "Price": 25.5, "Location": "Warehouse A, Shelf 2"}
XLSX_SHEET2_FIRST_ROW = {"Product ID": "P-300", "Description": "Wireless Mouse", "Price": 25.5, "Location": "Warehouse A, Shelf 2"}

XLS_SHEET1_FIRST_ROW = {"Index": 1.0, "First Name": "Dulce", "Last Name": "Abril", "Gender": "Female", "Country": "United States", "Age": 32.0}
XLS_SHEET2_FIRST_ROW = {"Index": 1.0, "Date": 43023.0, "Id": 1562.0}

XLS_SHEET1_HEADERS = ["Index", "First Name", "Last Name", "Gender", "Country", "Age"]
XLS_SHEET2_HEADERS = ["Index", "Date", "Id"]


# ---------------------------------------------------------------------------
# xlsx – shared across edit and stream modes
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("read_only", [False, True], ids=["xlsx_edit", "xlsx_stream"])
def test_xlsx_default_sheet_is_first_sheet(lib: RFExcelLibrary, read_only: bool):
    lib.load_workbook(XLSX_FILE, read_only=read_only)
    assert lib.get_rows()[0] == XLSX_SHEET1_FIRST_ROW


@pytest.mark.parametrize("read_only", [False, True], ids=["xlsx_edit", "xlsx_stream"])
def test_xlsx_switch_to_sheet2_changes_data(lib: RFExcelLibrary, read_only: bool):
    lib.load_workbook(XLSX_FILE, read_only=read_only)
    lib.switch_sheet("Sheet2")
    assert lib.get_rows()[0] == XLSX_SHEET2_FIRST_ROW


@pytest.mark.parametrize("read_only", [False, True], ids=["xlsx_edit", "xlsx_stream"])
def test_xlsx_switch_to_sheet2_correct_row_count(lib: RFExcelLibrary, read_only: bool):
    lib.load_workbook(XLSX_FILE, read_only=read_only)
    lib.switch_sheet("Sheet2")
    assert len(lib.get_rows()) == 4


# ---------------------------------------------------------------------------
# xlsx edit – unique tests
# ---------------------------------------------------------------------------

class TestSwitchSheetXlsxEdit:

    def test_switch_back_to_sheet1_restores_data(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.switch_sheet("Sheet2")
        lib.switch_sheet("List 1")
        assert lib.get_rows()[0] == XLSX_SHEET1_FIRST_ROW

    def test_switch_does_not_affect_sheet_list(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.switch_sheet("Sheet2")
        assert lib.list_sheet_names() == ["List 1", "Sheet2", "Sheet3", "Sheet4"]


# ---------------------------------------------------------------------------
# xlsx stream – unique tests
# ---------------------------------------------------------------------------

class TestSwitchSheetXlsxStream:

    def test_switch_resets_stream_position(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.get_rows()
        lib.switch_sheet("Sheet2")
        assert lib.get_rows()[0] == XLSX_SHEET2_FIRST_ROW

    def test_switch_back_to_sheet1_resets_stream(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.switch_sheet("Sheet2")
        lib.switch_sheet("List 1")
        assert lib.get_rows()[0] == XLSX_SHEET1_FIRST_ROW


# ---------------------------------------------------------------------------
# xls – shared across edit and on-demand modes
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("read_only", [False, True], ids=["xls_edit", "xls_on_demand"])
def test_xls_default_sheet_is_first_sheet(lib: RFExcelLibrary, read_only: bool):
    lib.load_workbook(XLS_FILE, read_only=read_only)
    assert lib.get_rows()[0] == XLS_SHEET1_FIRST_ROW


@pytest.mark.parametrize("read_only", [False, True], ids=["xls_edit", "xls_on_demand"])
def test_xls_switch_to_second_changes_data(lib: RFExcelLibrary, read_only: bool):
    lib.load_workbook(XLS_FILE, read_only=read_only)
    lib.switch_sheet("Second")
    assert lib.get_rows()[0] == XLS_SHEET2_FIRST_ROW


@pytest.mark.parametrize("read_only", [False, True], ids=["xls_edit", "xls_on_demand"])
def test_xls_switch_to_second_correct_row_count(lib: RFExcelLibrary, read_only: bool):
    lib.load_workbook(XLS_FILE, read_only=read_only)
    lib.switch_sheet("Second")
    assert len(lib.get_rows()) == 9


@pytest.mark.parametrize("read_only", [False, True], ids=["xls_edit", "xls_on_demand"])
def test_xls_switch_back_to_first_restores_data(lib: RFExcelLibrary, read_only: bool):
    lib.load_workbook(XLS_FILE, read_only=read_only)
    lib.switch_sheet("Second")
    lib.switch_sheet("First")
    assert lib.get_rows()[0] == XLS_SHEET1_FIRST_ROW


# ---------------------------------------------------------------------------
# xls edit – unique tests
# ---------------------------------------------------------------------------

class TestSwitchSheetXlsEdit:

    def test_switch_to_second_correct_headers(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        lib.switch_sheet("Second")
        rows = lib.get_rows()
        assert list(cast(dict[str, Any], rows[0]).keys()) == XLS_SHEET2_HEADERS

    def test_default_sheet_headers(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        rows = lib.get_rows()
        assert list(cast(dict[str, Any], rows[0]).keys()) == XLS_SHEET1_HEADERS


# ---------------------------------------------------------------------------
# csv
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

    @pytest.mark.parametrize(
        ("path", "read_only"),
        [
            (XLSX_FILE, False),
            (XLSX_FILE, True),
            (XLS_FILE,  False),
            (XLS_FILE,  True),
        ],
        ids=["xlsx_edit", "xlsx_stream", "xls_edit", "xls_on_demand"],
    )
    def test_switch_to_nonexistent_sheet_raises(
        self, lib: RFExcelLibrary, path: str, read_only: bool
    ):
        lib.load_workbook(path, read_only=read_only)
        with pytest.raises(SheetDoesNotExistException, match="DoesNotExist"):
            lib.switch_sheet("DoesNotExist")

    def test_switch_to_nonexistent_sheet_csv_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE)
        with pytest.raises(OperationNotSupportedForFormat):
            lib.switch_sheet("DoesNotExist")
