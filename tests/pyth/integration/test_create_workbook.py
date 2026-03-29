from pathlib import Path

import pytest

from rfexcel.exception.library_exceptions import (
    FileAlreadyExistsException, FileFormatNotSupportedException,
    HeadersNotDeterminedException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLSX_FILE

# ---------------------------------------------------------------------------
# positive
# ---------------------------------------------------------------------------

class TestCreateWorkbookPositive:

    @pytest.mark.parametrize("filename", ["new.xlsx", "new.csv"], ids=["xlsx", "csv"])
    def test_create_sets_active_workbook(
        self, lib: RFExcelLibrary, tmp_path: Path, filename: str
    ):
        lib.create_workbook(str(tmp_path / filename))
        assert lib._active_workbook is not None

    @pytest.mark.parametrize("filename", ["new.xlsx", "new.csv"], ids=["xlsx", "csv"])
    def test_create_produces_file_on_disk(
        self, lib: RFExcelLibrary, tmp_path: Path, filename: str
    ):
        path = tmp_path / filename
        lib.create_workbook(str(path))
        assert path.exists()

    def test_created_xlsx_is_immediately_readable(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(tmp_path / "empty.xlsx")
        lib.create_workbook(path)
        rows = lib.get_rows()
        assert rows == []

    def test_created_csv_get_rows_raises_on_empty_file(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(tmp_path / "empty.csv")
        lib.create_workbook(path)
        with pytest.raises(HeadersNotDeterminedException):
            lib.get_rows()

# ---------------------------------------------------------------------------
# negative
# ---------------------------------------------------------------------------

class TestCreateWorkbookNegative:

    @pytest.mark.parametrize("filename", ["existing.xlsx", "existing.csv"], ids=["xlsx", "csv"])
    def test_create_on_existing_file_raises(
        self, lib: RFExcelLibrary, tmp_path: Path, filename: str
    ):
        path = str(tmp_path / filename)
        lib.create_workbook(path)
        lib.close()
        with pytest.raises(FileAlreadyExistsException):
            lib.create_workbook(path)

    @pytest.mark.parametrize(
        "filename",
        ["legacy.xls", "notes.txt", "sheet.ods"],
        ids=["xls", "txt", "ods"],
    )
    def test_create_unsupported_format_raises(
        self, lib: RFExcelLibrary, tmp_path: Path, filename: str
    ):
        with pytest.raises(FileFormatNotSupportedException):
            lib.create_workbook(str(tmp_path / filename))

    def test_active_workbook_changed_after_failed_create(self, lib: RFExcelLibrary, tmp_path: Path):
        lib.load_workbook(XLSX_FILE)
        active_before = lib._active_workbook
        with pytest.raises(FileFormatNotSupportedException):
            lib.create_workbook(str(tmp_path / "bad.txt"))
        assert lib._active_workbook is not active_before

# ---------------------------------------------------------------------------
# edge cases
# ---------------------------------------------------------------------------

class TestCreateWorkbookEdge:

    def test_create_xlsx_with_nested_new_directories(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(tmp_path / "a" / "b" / "c" / "deep.xlsx")
        lib.create_workbook(path)
        assert Path(path).exists()

    @pytest.mark.parametrize("filename", ["roundtrip.xlsx", "roundtrip.csv"], ids=["xlsx", "csv"])
    def test_created_file_can_be_loaded_afterwards(
        self, lib: RFExcelLibrary, tmp_path: Path, filename: str
    ):
        path = str(tmp_path / filename)
        lib.create_workbook(path)
        lib.load_workbook(path)
        assert lib._active_workbook is not None

    def test_two_different_workbooks_created_independently(self, lib: RFExcelLibrary, tmp_path: Path):
        path_a = str(tmp_path / "a.xlsx")
        path_b = str(tmp_path / "b.xlsx")
        lib.create_workbook(path_a)
        lib.create_workbook(path_b)
        assert Path(path_a).exists()
        assert Path(path_b).exists()
