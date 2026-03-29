import shutil
from pathlib import Path

import pytest

from rfexcel.exception.library_exceptions import (LibraryException,
                                                  NullComponentException,
                                                  RowIndexOutOfBoundsException,
                                                  WorkbookNotOpenException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

# ---------------------------------------------------------------------------
# XLSX – Edit mode
# ---------------------------------------------------------------------------

class TestDeleteRowXlsxEdit:

    def test_row_is_removed(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.delete_row(2)
        rows = lib.get_rows()
        assert all(r["Product ID"] != "P-200" for r in rows)

    def test_row_count_decreases(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        before = len(lib.get_rows())
        lib.delete_row(2)
        assert len(lib.get_rows()) == before - 1

    def test_remaining_rows_readable_in_order(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.delete_row(3)
        ids = [r["Product ID"] for r in lib.get_rows()]
        assert "P-200" in ids
        assert "P-201" not in ids
        assert "P-202" in ids
        assert "P-203" in ids

    def test_delete_last_data_row(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        before = len(lib.get_rows())
        lib.delete_row(5)
        rows = lib.get_rows()
        assert len(rows) == before - 1
        assert all(r["Product ID"] != "P-203" for r in rows)

    def test_row_number_zero_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(RowIndexOutOfBoundsException):
            lib.delete_row(0)

    def test_row_number_negative_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(RowIndexOutOfBoundsException):
            lib.delete_row(-1)

    def test_row_number_beyond_last_row_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(RowIndexOutOfBoundsException):
            lib.delete_row(9999)

    def test_delete_header_row_removes_first_row(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        before = lib.get_row(2)
        lib.delete_row(1)
        assert lib.get_row(1) == before


# ---------------------------------------------------------------------------
# Read-only / streaming modes – raises for all formats
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    ("path", "read_only"),
    [
        (XLSX_FILE, True),
        (XLS_FILE,  True),
        (CSV_FILE,  True),
    ],
    ids=["xlsx_stream", "xls_on_demand", "csv_stream"],
)
def test_raises_in_read_only_mode(lib: RFExcelLibrary, path: str, read_only: bool):
    lib.load_workbook(path, read_only=read_only)
    with pytest.raises(NullComponentException):
        lib.delete_row(2)


# ---------------------------------------------------------------------------
# XLS – Edit mode (lazy conversion)
# ---------------------------------------------------------------------------

class TestDeleteRowXlsEdit:

    def test_delete_triggers_conversion(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        before = len(lib.get_rows())
        lib.delete_row(2)
        rows = lib.get_rows()
        assert len(rows) == before - 1


# ---------------------------------------------------------------------------
# CSV – Edit mode
# ---------------------------------------------------------------------------

class TestDeleteRowCsvEdit:

    def test_row_deleted(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        lib.delete_row(2)
        assert all(r["Product ID"] != "P-200" for r in lib.get_rows())

    def test_row_count_decreases(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        before = len(lib.get_rows())
        lib.delete_row(2)
        assert len(lib.get_rows()) == before - 1

    def test_row_number_beyond_last_row_raises(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        with pytest.raises(RowIndexOutOfBoundsException):
            lib.delete_row(9999)


# ---------------------------------------------------------------------------
# No workbook open
# ---------------------------------------------------------------------------

class TestDeleteRowNoWorkbook:

    def test_raises_when_no_workbook_open(self, lib: RFExcelLibrary):
        with pytest.raises(WorkbookNotOpenException):
            lib.delete_row(2)
