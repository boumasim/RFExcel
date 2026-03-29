import shutil
from pathlib import Path
from typing import Any, cast

import pytest

from rfexcel.exception.library_exceptions import (
    HeadersNotDeterminedException, NullComponentException,
    WorkbookNotOpenException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

# ---------------------------------------------------------------------------
# XLSX – Edit mode
# ---------------------------------------------------------------------------

class TestDeleteRowsXlsxEdit:

    def test_returns_count_of_deleted_rows(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        count = lib.delete_rows(search_criteria={"Product ID": "P-200"})
        assert count == 1

    def test_row_is_removed(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.delete_rows(search_criteria={"Product ID": "P-200"})
        rows = cast(list[dict[str, Any]], lib.get_rows())
        assert all(r["Product ID"] != "P-200" for r in rows)

    def test_row_count_decreases(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        before = len(lib.get_rows())
        lib.delete_rows(search_criteria={"Product ID": "P-201"})
        assert len(lib.get_rows()) == before - 1

    def test_no_match_returns_zero(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        before = lib.get_rows()
        count = lib.delete_rows(search_criteria={"Product ID": "NONEXISTENT"})
        assert count == 0
        assert lib.get_rows() == before

    def test_multiple_matches_all_deleted(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.update_values(search_criteria={"Product ID": "P-201"}, values={"Location": "SAME"})
        lib.update_values(search_criteria={"Product ID": "P-202"}, values={"Location": "SAME"})
        count = lib.delete_rows(search_criteria={"Location": "SAME"})
        assert count == 2
        rows = cast(list[dict[str, Any]], lib.get_rows())
        assert all(r["Location"] != "SAME" for r in rows)

    def test_first_only_deletes_single_row(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.update_values(search_criteria={"Product ID": "P-201"}, values={"Location": "SAME"})
        lib.update_values(search_criteria={"Product ID": "P-202"}, values={"Location": "SAME"})
        count = lib.delete_rows(search_criteria={"Location": "SAME"}, one_row=True)
        assert count == 1
        rows = lib.get_rows()
        assert sum(1 for r in rows if r["Location"] == "SAME") == 1

    def test_partial_match_deletes_substring_rows(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        before = len(lib.get_rows())
        count = lib.delete_rows(search_criteria={"Location": "Warehouse"}, partial_match=True)
        assert count >= 1
        assert len(lib.get_rows()) == before - count
        assert all("Warehouse" not in r["Location"] for r in lib.get_rows())

    def test_dict_search_criteria_numeric_value_matches_native_type(self, lib: RFExcelLibrary):
        """Delete with numeric value matches XLSX native int type."""
        lib.load_workbook(XLSX_FILE)
        count = lib.delete_rows(search_criteria={"Price": "150"})
        assert count == 1
        rows = cast(list[dict[str, Any]], lib.get_rows())
        assert all(r["Product ID"] != "P-202" for r in rows)

    def test_remaining_rows_are_readable(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.delete_rows(search_criteria={"Product ID": "P-202"})
        ids = [r["Product ID"] for r in cast(list[dict[str, Any]], lib.get_rows())]
        assert "P-200" in ids
        assert "P-201" in ids
        assert "P-202" not in ids
        assert "P-203" in ids

    def test_header_row_out_of_range_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(HeadersNotDeterminedException):
            lib.delete_rows(search_criteria={"Product ID": "P-200"}, header_row=9999)


# ---------------------------------------------------------------------------
# Read-only / streaming modes – raises for all formats
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    ("path", "criteria"),
    [
        (XLSX_FILE, {"Product ID": "P-200"}),
        (XLS_FILE,  {"First Name": "Dulce"}),
        (CSV_FILE,  {"Product ID": "P-200"}),
    ],
    ids=["xlsx_stream", "xls_on_demand", "csv_stream"],
)
def test_raises_in_read_only_mode(lib: RFExcelLibrary, path: str, criteria: dict):
    lib.load_workbook(path, read_only=True)
    with pytest.raises(NullComponentException):
        lib.delete_rows(search_criteria=criteria)


# ---------------------------------------------------------------------------
# XLS – Edit mode
# ---------------------------------------------------------------------------

class TestDeleteRowsXlsEdit:

    def test_delete_triggers_conversion(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        before = len(lib.get_rows())
        count = lib.delete_rows(search_criteria={"First Name": "Dulce"})
        assert count == 1
        assert len(lib.get_rows()) == before - 1
        rows = cast(list[dict[str, Any]], lib.get_rows())
        assert all(r["First Name"] != "Dulce" for r in rows)


# ---------------------------------------------------------------------------
# CSV – Edit mode
# ---------------------------------------------------------------------------

class TestDeleteRowsCsvEdit:

    def test_matched_row_deleted(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        count = lib.delete_rows(search_criteria={"Product ID": "P-200"})
        assert count == 1
        rows = cast(list[dict[str, Any]], lib.get_rows())
        assert all(r["Product ID"] != "P-200" for r in rows)

    def test_row_count_decreases(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        before = len(lib.get_rows())
        lib.delete_rows(search_criteria={"Product ID": "P-201"})
        assert len(lib.get_rows()) == before - 1

    def test_first_only_deletes_single_row(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        lib.update_values(search_criteria={"Product ID": "P-200"}, values={"Location": "DUPLICATE_LOC"})
        lib.update_values(search_criteria={"Product ID": "P-201"}, values={"Location": "DUPLICATE_LOC"})
        count = lib.delete_rows(search_criteria={"Location": "DUPLICATE_LOC"}, one_row=True)
        assert count == 1
        rows = lib.get_rows()
        assert sum(1 for r in rows if r["Location"] == "DUPLICATE_LOC") == 1


# ---------------------------------------------------------------------------
# No workbook open
# ---------------------------------------------------------------------------

class TestDeleteRowsNoWorkbook:

    def test_raises_when_no_workbook_open(self, lib: RFExcelLibrary):
        with pytest.raises(WorkbookNotOpenException):
            lib.delete_rows(search_criteria={"Product ID": "P-200"})
