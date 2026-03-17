"""Integration tests for the Append Rows keyword."""
import shutil

import pytest

from rfexcel.exception.library_exceptions import LibraryException
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

_ROW_A = {"Product ID": "P-001", "Description": "Alpha", "Price": "1.00", "Location": "Shelf-A"}
_ROW_B = {"Product ID": "P-002", "Description": "Beta",  "Price": "2.00", "Location": "Shelf-B"}


# ---------------------------------------------------------------------------
# XLSX – Edit mode
# ---------------------------------------------------------------------------

class TestAppendRowsXlsxEdit:

    def test_all_rows_appended_in_order(self, lib: RFExcelLibrary, tmp_path):
        path = str(shutil.copy(XLSX_FILE, tmp_path / "data.xlsx"))
        lib.load_workbook(path)
        before = len(lib.get_rows())
        lib.append_rows([_ROW_A, _ROW_B])
        rows = lib.get_rows()
        assert len(rows) == before + 2
        assert rows[-2]["Product ID"] == "P-001"
        assert rows[-1]["Product ID"] == "P-002"

    def test_empty_list_is_noop(self, lib: RFExcelLibrary, tmp_path):
        path = str(shutil.copy(XLSX_FILE, tmp_path / "data.xlsx"))
        lib.load_workbook(path)
        before = len(lib.get_rows())
        lib.append_rows([])
        assert len(lib.get_rows()) == before

    def test_partial_rows_fill_missing_with_empty_string(self, lib: RFExcelLibrary, tmp_path):
        path = str(shutil.copy(XLSX_FILE, tmp_path / "data.xlsx"))
        lib.load_workbook(path)
        lib.append_rows([{"Product ID": "P-010"}, {"Price": "5.00"}])
        rows = lib.get_rows()
        assert rows[-2]["Product ID"] == "P-010"
        assert rows[-2]["Description"] == ""
        assert rows[-1]["Price"] == "5.00"
        assert rows[-1]["Product ID"] == ""

    def test_rows_persisted_after_save(self, lib: RFExcelLibrary, tmp_path):
        path = str(shutil.copy(XLSX_FILE, tmp_path / "data.xlsx"))
        lib.load_workbook(path)
        lib.append_rows([_ROW_A, _ROW_B])
        lib.save_workbook()
        lib.close()

        lib2 = RFExcelLibrary()
        lib2.load_workbook(path)
        ids = [r["Product ID"] for r in lib2.get_rows()]
        assert "P-001" in ids
        assert "P-002" in ids
        lib2.close()


# ---------------------------------------------------------------------------
# XLSX – Streaming mode
# ---------------------------------------------------------------------------

class TestAppendRowsXlsxStream:

    def test_raises_in_stream_mode(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE, read_only=True)
        with pytest.raises(LibraryException):
            lib.append_rows([_ROW_A])


# ---------------------------------------------------------------------------
# XLS – Edit mode (lazy conversion)
# ---------------------------------------------------------------------------

class TestAppendRowsXlsEdit:

    _XLS_ROW_A = {"Index": "99", "First Name": "Alice", "Last Name": "Smith",
                  "Gender": "Female", "Country": "Czech Republic", "Age": "30"}
    _XLS_ROW_B = {"Index": "100", "First Name": "Bob", "Last Name": "Jones",
                  "Gender": "Male", "Country": "Slovakia", "Age": "25"}

    def test_rows_appended_after_lazy_conversion(self, lib: RFExcelLibrary, tmp_path):
        path = str(shutil.copy(XLS_FILE, tmp_path / "example.xls"))
        lib.load_workbook(path)
        before = len(lib.get_rows())
        lib.append_rows([self._XLS_ROW_A, self._XLS_ROW_B])
        rows = lib.get_rows()
        assert len(rows) == before + 2
        assert rows[-2]["Index"] == "99"
        assert rows[-1]["Index"] == "100"


# ---------------------------------------------------------------------------
# CSV – Edit mode
# ---------------------------------------------------------------------------

class TestAppendRowsCsvEdit:

    def test_rows_appended_and_read_back(self, lib: RFExcelLibrary, tmp_path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        before = len(lib.get_rows())
        lib.append_rows([_ROW_A, _ROW_B])
        lib.save_workbook()
        lib.close()

        lib2 = RFExcelLibrary()
        lib2.load_workbook(path)
        rows = lib2.get_rows()
        assert len(rows) == before + 2
        assert rows[-2]["Product ID"] == "P-001"
        assert rows[-1]["Product ID"] == "P-002"
        lib2.close()

    def test_raises_in_stream_mode(self, lib: RFExcelLibrary):
        lib.load_workbook(CSV_FILE, read_only=True)
        with pytest.raises(LibraryException):
            lib.append_rows([_ROW_A])
