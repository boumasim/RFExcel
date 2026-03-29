import shutil
from pathlib import Path
from typing import Any, cast

import pytest

from rfexcel.exception.library_exceptions import (
    HeadersNotDeterminedException, NullComponentException,
    WorkbookNotOpenException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

_XLSX_HEADER_ROW = 1


# ---------------------------------------------------------------------------
# XLSX – Edit mode
# ---------------------------------------------------------------------------

class TestUpdateValuesXlsxEdit:

    def test_returns_count_of_updated_rows(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        count = lib.update_values(
            search_criteria={"Product ID": "P-200"},
            values={"Price": "0.00"},
            header_row=_XLSX_HEADER_ROW,
        )
        assert count == 1

    def test_matching_column_is_changed(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.update_values(
            search_criteria={"Product ID": "P-200"},
            values={"Price": "0.00"},
            header_row=_XLSX_HEADER_ROW,
        )
        rows = lib.get_rows(header_row=_XLSX_HEADER_ROW)
        row = next(r for r in rows if r["Product ID"] == "P-200")
        assert row["Price"] == "0.00"

    def test_unspecified_columns_are_not_touched(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows_before = cast(list[dict[str, Any]], lib.get_rows(header_row=_XLSX_HEADER_ROW))
        original_desc = next(r["Description"] for r in rows_before if r["Product ID"] == "P-201")

        lib.update_values(
            search_criteria={"Product ID": "P-201"},
            values={"Price": "1.00"},
            header_row=_XLSX_HEADER_ROW,
        )
        rows = lib.get_rows(header_row=_XLSX_HEADER_ROW)
        row = next(r for r in rows if r["Product ID"] == "P-201")
        assert row["Description"] == original_desc
        assert row["Price"] == "1.00"

    def test_no_match_returns_zero_and_leaves_data_unchanged(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows_before = lib.get_rows(header_row=_XLSX_HEADER_ROW)
        count = lib.update_values(
            search_criteria={"Product ID": "NONEXISTENT"},
            values={"Price": "999.99"},
            header_row=_XLSX_HEADER_ROW,
        )
        assert count == 0
        assert lib.get_rows(header_row=_XLSX_HEADER_ROW) == rows_before

    def test_multiple_matching_rows_all_updated(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.update_values(
            search_criteria={"Product ID": "P-201"},
            values={"Location": "SAME"},
            header_row=_XLSX_HEADER_ROW,
        )
        lib.update_values(
            search_criteria={"Product ID": "P-202"},
            values={"Location": "SAME"},
            header_row=_XLSX_HEADER_ROW,
        )
        count = lib.update_values(
            search_criteria={"Location": "SAME"},
            values={"Price": "0.00"},
            header_row=_XLSX_HEADER_ROW,
        )
        assert count == 2
        rows = lib.get_rows(header_row=_XLSX_HEADER_ROW)
        for r in rows:
            if r["Location"] == "SAME":
                assert r["Price"] == "0.00"

    def test_partial_match_updates_substring_rows(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        count = lib.update_values(
            search_criteria={"Location": "Warehouse"},
            values={"Price": "0.01"},
            header_row=_XLSX_HEADER_ROW,
            partial_match=True,
        )
        assert count >= 1
        rows = cast(list[dict[str, Any]], lib.get_rows(header_row=_XLSX_HEADER_ROW))
        updated = [r for r in rows if "Warehouse" in r["Location"]]
        assert all(r["Price"] == "0.01" for r in updated)

    def test_string_search_criteria_accepted(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        count = lib.update_values(
            search_criteria="Product ID=P-203",
            values={"Location": "Archived"},
            header_row=_XLSX_HEADER_ROW,
        )
        assert count == 1
        rows = lib.get_rows(header_row=_XLSX_HEADER_ROW)
        row = next(r for r in rows if r["Product ID"] == "P-203")
        assert row["Location"] == "Archived"

    def test_dict_search_criteria_numeric_value_matches_native_type(self, lib: RFExcelLibrary):
        """Search with numeric value matches XLSX native int type."""
        lib.load_workbook(XLSX_FILE)
        count = lib.update_values(
            search_criteria={"Price": "150"},
            values={"Location": "Updated"},
            header_row=_XLSX_HEADER_ROW,
        )
        assert count == 1
        rows = lib.get_rows(header_row=_XLSX_HEADER_ROW)
        row = next(r for r in rows if r["Product ID"] == "P-202")
        assert row["Location"] == "Updated"

    def test_header_row_out_of_range_raises(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        with pytest.raises(HeadersNotDeterminedException):
            lib.update_values(
                search_criteria={"Product ID": "P-200"},
                values={"Price": "0.00"},
                header_row=9999,
            )

    def test_unknown_value_keys_silently_ignored(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        rows_before = lib.get_rows(header_row=_XLSX_HEADER_ROW)
        count = lib.update_values(
            search_criteria={"Product ID": "P-200"},
            values={"NonExistentColumn": "X"},
            header_row=_XLSX_HEADER_ROW,
        )
        assert count == 1
        assert lib.get_rows(header_row=_XLSX_HEADER_ROW) == rows_before


# ---------------------------------------------------------------------------
# Read-only / streaming modes – raises for all formats
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    ("path", "criteria", "vals"),
    [
        (XLSX_FILE, {"Product ID": "P-200"}, {"Price": "0.00"}),
        (XLS_FILE,  {"First Name": "Dulce"},  {"Country": "Updated"}),
        (CSV_FILE,  {"Product ID": "P-200"}, {"Price": "0.00"}),
    ],
    ids=["xlsx_stream", "xls_on_demand", "csv_stream"],
)
def test_raises_in_read_only_mode(
    lib: RFExcelLibrary, path: str, criteria: dict, vals: dict
):
    lib.load_workbook(path, read_only=True)
    with pytest.raises(NullComponentException):
        lib.update_values(search_criteria=criteria, values=vals)


# ---------------------------------------------------------------------------
# XLS – Edit mode
# ---------------------------------------------------------------------------

class TestUpdateValuesXlsEdit:

    def test_update_triggers_conversion(self, lib: RFExcelLibrary):
        lib.load_workbook(XLS_FILE)
        count = lib.update_values(
            search_criteria={"First Name": "Dulce"},
            values={"Country": "Updated"},
        )
        assert count == 1
        row = next(r for r in lib.get_rows() if r["First Name"] == "Dulce")
        assert row["Country"] == "Updated"


# ---------------------------------------------------------------------------
# CSV – Edit mode
# ---------------------------------------------------------------------------

class TestUpdateValuesCsvEdit:

    def test_matching_row_updated(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        count = lib.update_values(
            search_criteria={"Product ID": "P-200"},
            values={"Price": "0.00"},
        )
        assert count == 1
        row = next(r for r in lib.get_rows() if r["Product ID"] == "P-200")
        assert row["Price"] == 0

    def test_unspecified_columns_untouched(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        original_desc = next(
            r["Description"] for r in cast(list[dict[str, Any]], lib.get_rows()) if r["Product ID"] == "P-201"
        )
        lib.update_values(
            search_criteria={"Product ID": "P-201"},
            values={"Price": "1.11"},
        )
        row = next(r for r in lib.get_rows() if r["Product ID"] == "P-201")
        assert row["Description"] == original_desc
        assert row["Price"] == 1.11

    def test_partial_match(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        count = lib.update_values(
            search_criteria={"Location": "Online"},
            values={"Price": "FREE"},
            partial_match=True,
        )
        assert count >= 1
        rows = cast(list[dict[str, Any]], lib.get_rows())
        assert all(
            r["Price"] == "FREE" for r in rows if "Online" in r["Location"]
        )


# ---------------------------------------------------------------------------
# No workbook open
# ---------------------------------------------------------------------------

class TestUpdateValuesNoWorkbook:

    def test_raises_when_no_workbook_open(self, lib: RFExcelLibrary):
        with pytest.raises(WorkbookNotOpenException):
            lib.update_values(
                search_criteria={"Product ID": "P-200"},
                values={"Price": "0.00"},
            )


# ---------------------------------------------------------------------------
# Update Values with first_only=True
# ---------------------------------------------------------------------------

class TestUpdateFirst:

    def test_returns_1_when_match_found(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        count = lib.update_values(
            search_criteria={"Product ID": "P-200"},
            values={"Price": "0.01"},
            first_only=True,
        )
        assert count == 1

    def test_returns_0_when_no_match(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        count = lib.update_values(
            search_criteria={"Product ID": "NONEXISTENT"},
            values={"Price": "0.01"},
            first_only=True,
        )
        assert count == 0

    def test_only_first_match_is_updated(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.update_values(search_criteria={"Product ID": "P-201"}, values={"Location": "SHARED"})
        lib.update_values(search_criteria={"Product ID": "P-202"}, values={"Location": "SHARED"})

        lib.update_values(
            search_criteria={"Location": "SHARED"},
            values={"Price": "FIRST_ONLY"},
            first_only=True,
        )
        rows = lib.get_rows()
        updated = [r for r in rows if r["Location"] == "SHARED"]
        assert len(updated) == 2
        assert sum(1 for r in updated if r["Price"] == "FIRST_ONLY") == 1

    def test_correct_column_changed(self, lib: RFExcelLibrary):
        lib.load_workbook(XLSX_FILE)
        lib.update_values(
            search_criteria={"Product ID": "P-203"},
            values={"Location": "Updated"},
            first_only=True,
        )
        rows = lib.get_rows()
        row = next(r for r in rows if r["Product ID"] == "P-203")
        assert row["Location"] == "Updated"

    def test_returns_0_when_no_workbook_open(self, lib: RFExcelLibrary):
        with pytest.raises(WorkbookNotOpenException):
            lib.update_values(
                search_criteria={"Product ID": "P-200"},
                values={"Price": "0.00"},
                first_only=True,
            )

    def test_csv_only_first_match_updated(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        count = lib.update_values(
            search_criteria={"Location": "Online"},
            values={"Price": "FIRST"},
            partial_match=True,
            first_only=True,
        )
        assert count == 1
        rows = lib.get_rows()
        updated = [r for r in rows if r["Price"] == "FIRST"]
        assert len(updated) == 1
