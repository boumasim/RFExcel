import shutil
from pathlib import Path

import pytest

from rfexcel.exception.library_exceptions import (FileSaveException,
                                                  NullComponentException,
                                                  WorkbookNotOpenException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

# ---------------------------------------------------------------------------
# XLSX – Edit mode
# ---------------------------------------------------------------------------

class TestSaveWorkbookXlsxEdit:

    def test_save_in_place_persists_changes(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(XLSX_FILE, tmp_path / "data.xlsx"))
        lib.load_workbook(path)
        lib.add_sheet("Persisted")
        lib.save_workbook()
        lib.close()

        lib2 = RFExcelLibrary()
        lib2.load_workbook(path)
        assert "Persisted" in lib2.list_sheet_names()
        lib2.close()

    def test_save_as_creates_new_file(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(XLSX_FILE, tmp_path / "data.xlsx"))
        new_path = tmp_path / "copy.xlsx"
        lib.load_workbook(path)
        lib.save_workbook(str(new_path))
        assert new_path.exists()

    def test_save_as_does_not_modify_original(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(XLSX_FILE, tmp_path / "data.xlsx"))
        new_path = str(tmp_path / "copy.xlsx")
        lib.load_workbook(path)
        lib.add_sheet("OnlyInCopy")
        lib.save_workbook(new_path)
        lib.close()

        lib2 = RFExcelLibrary()
        lib2.load_workbook(path)
        assert "OnlyInCopy" not in lib2.list_sheet_names()
        lib2.close()

    def test_save_as_updates_active_path(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(XLSX_FILE, tmp_path / "data.xlsx"))
        new_path = str(tmp_path / "moved.xlsx")
        lib.load_workbook(path)
        lib.save_workbook(new_path)
        lib.add_sheet("SecondSave")
        lib.save_workbook()
        lib.close()

        lib2 = RFExcelLibrary()
        lib2.load_workbook(new_path)
        assert "SecondSave" in lib2.list_sheet_names()
        lib2.close()

    def test_save_preserves_all_existing_sheets(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(XLSX_FILE, tmp_path / "data.xlsx"))
        lib.load_workbook(path)
        names_before = lib.list_sheet_names()
        lib.save_workbook()
        lib.close()

        lib2 = RFExcelLibrary()
        lib2.load_workbook(path)
        assert lib2.list_sheet_names() == names_before
        lib2.close()


# ---------------------------------------------------------------------------
# Read-only / streaming modes – raises for all formats
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    ("source", "filename"),
    [
        (XLSX_FILE, "data.xlsx"),
        (XLS_FILE,  "example.xls"),
        (CSV_FILE,  "data.csv"),
    ],
    ids=["xlsx_stream", "xls_on_demand", "csv_stream"],
)
def test_save_raises_in_read_only_mode(
    lib: RFExcelLibrary, tmp_path: Path, source: str, filename: str
):
    path = str(shutil.copy(source, tmp_path / filename))
    lib.load_workbook(path, read_only=True)
    with pytest.raises(NullComponentException):
        lib.save_workbook()


# ---------------------------------------------------------------------------
# XLS – Edit mode
# ---------------------------------------------------------------------------

class TestSaveWorkbookXlsEdit:

    def test_save_triggers_implicit_conversion_and_produces_file(
        self, lib: RFExcelLibrary, tmp_path: Path
    ):
        path = str(shutil.copy(XLS_FILE, tmp_path / "example.xls"))
        new_path = str(tmp_path / "result.xlsx")
        lib.load_workbook(path)
        lib.save_workbook(new_path)
        lib.close()
        assert (tmp_path / "result.xlsx").exists()

    def test_save_as_xlsx_succeeds_without_prior_write_op(
        self, lib: RFExcelLibrary, tmp_path: Path
    ):
        path = str(shutil.copy(XLS_FILE, tmp_path / "example.xls"))
        new_path = tmp_path / "converted.xlsx"
        lib.load_workbook(path)
        lib.save_workbook(str(new_path))
        lib.close()

        lib2 = RFExcelLibrary()
        lib2.load_workbook(str(new_path))
        assert isinstance(lib2.list_sheet_names(), list)
        lib2.close()

    def test_save_preserves_added_sheet(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(XLS_FILE, tmp_path / "example.xls"))
        new_path = str(tmp_path / "out.xlsx")
        lib.load_workbook(path)
        lib.add_sheet("NewSheet")
        lib.save_workbook(new_path)
        lib.close()

        lib2 = RFExcelLibrary()
        lib2.load_workbook(new_path)
        assert "NewSheet" in lib2.list_sheet_names()
        lib2.close()

    def test_original_xls_untouched_after_save(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(XLS_FILE, tmp_path / "example.xls"))
        new_path = str(tmp_path / "out.xlsx")
        lib.load_workbook(path)
        lib.add_sheet("NewSheet")
        lib.save_workbook(new_path)
        lib.close()

        lib2 = RFExcelLibrary()
        lib2.load_workbook(path)
        assert "NewSheet" not in lib2.list_sheet_names()
        lib2.close()


# ---------------------------------------------------------------------------
# CSV – Edit mode
# ---------------------------------------------------------------------------

class TestSaveWorkbookCsvEdit:

    def test_save_in_place_produces_readable_file(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        lib.load_workbook(path)
        lib.save_workbook()
        lib.close()

        lib2 = RFExcelLibrary()
        lib2.load_workbook(path)
        assert isinstance(lib2.get_rows(), list)
        lib2.close()

    def test_save_as_creates_new_csv_file(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        new_path = tmp_path / "output.csv"
        lib.load_workbook(path)
        lib.save_workbook(str(new_path))
        assert new_path.exists()

    def test_save_as_preserves_content(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(shutil.copy(CSV_FILE, tmp_path / "data.csv"))
        new_path = str(tmp_path / "output.csv")
        lib.load_workbook(path)
        rows_original = lib.get_rows()
        lib.save_workbook(new_path)
        lib.close()

        lib2 = RFExcelLibrary()
        lib2.load_workbook(new_path)
        assert lib2.get_rows() == rows_original
        lib2.close()


# ---------------------------------------------------------------------------
# No workbook open
# ---------------------------------------------------------------------------

class TestSaveWorkbookNoWorkbook:

    def test_save_raises_when_no_workbook_open(self, lib: RFExcelLibrary):
        with pytest.raises(WorkbookNotOpenException):
            lib.save_workbook()


# ---------------------------------------------------------------------------
# FileSaveException – bad path
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    ("source", "out_filename"),
    [
        (XLSX_FILE, "out.xlsx"),
        (CSV_FILE,  "out.csv"),
    ],
    ids=["xlsx", "csv"],
)
def test_save_to_nonexistent_dir_raises(
    lib: RFExcelLibrary, tmp_path: Path, source: str, out_filename: str
):
    path = str(shutil.copy(source, tmp_path / Path(source).name))
    lib.load_workbook(path)
    with pytest.raises(FileSaveException):
        lib.save_workbook(str(tmp_path / "no_such_dir" / out_filename))
