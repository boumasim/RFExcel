"""Integration tests for the Create Workbook keyword.

Every test that actually creates a file uses pytest's tmp_path fixture so no
test artifact is left in the source tree and tests are fully isolated.

Covers:
  - Positive: create xlsx, create csv; file exists on disk; is immediately writable.
  - Negative: file already exists, unsupported formats (.xls, .txt).
  - Edge: parent directories created automatically, created file is loadable.
"""
from pathlib import Path

import pytest

from rfexcel.exception.library_exceptions import (
    FileAlreadyExistsException, FileFormatNotSupportedException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLSX_FILE

# ─── positive ─────────────────────────────────────────────────────────────────

class TestCreateWorkbookPositive:

    def test_create_xlsx_sets_active_workbook(self, lib: RFExcelLibrary, tmp_path):
        path = str(tmp_path / "new.xlsx")
        lib.create_workbook(path)
        assert lib._active_workbook is not None

    def test_create_xlsx_produces_file_on_disk(self, lib: RFExcelLibrary, tmp_path):
        path = tmp_path / "new.xlsx"
        lib.create_workbook(str(path))
        assert path.exists()

    def test_create_csv_sets_active_workbook(self, lib: RFExcelLibrary, tmp_path):
        path = str(tmp_path / "new.csv")
        lib.create_workbook(path)
        assert lib._active_workbook is not None

    def test_create_csv_produces_file_on_disk(self, lib: RFExcelLibrary, tmp_path):
        path = tmp_path / "new.csv"
        lib.create_workbook(str(path))
        assert path.exists()

    def test_created_xlsx_is_immediately_readable(self, lib: RFExcelLibrary, tmp_path):
        """A freshly created (empty) xlsx should return an empty row list."""
        path = str(tmp_path / "empty.xlsx")
        lib.create_workbook(path)
        rows = lib.get_rows()
        assert rows == []

    def test_created_csv_is_immediately_readable(self, lib: RFExcelLibrary, tmp_path):
        path = str(tmp_path / "empty.csv")
        lib.create_workbook(path)
        rows = lib.get_rows()
        assert rows == []


# ─── negative ─────────────────────────────────────────────────────────────────

class TestCreateWorkbookNegative:

    def test_create_on_existing_xlsx_raises(self, lib: RFExcelLibrary, tmp_path):
        """Creating a workbook where a file already exists must raise."""
        path = str(tmp_path / "existing.xlsx")
        lib.create_workbook(path)
        lib.close()
        with pytest.raises(FileAlreadyExistsException):
            lib.create_workbook(path)

    def test_create_on_existing_csv_raises(self, lib: RFExcelLibrary, tmp_path):
        path = str(tmp_path / "existing.csv")
        lib.create_workbook(path)
        lib.close()
        with pytest.raises(FileAlreadyExistsException):
            lib.create_workbook(path)

    def test_create_xls_raises_format_not_supported(self, lib: RFExcelLibrary, tmp_path):
        """Writing .xls is not supported — only reading is."""
        path = str(tmp_path / "legacy.xls")
        with pytest.raises(FileFormatNotSupportedException):
            lib.create_workbook(path)

    def test_create_txt_raises_format_not_supported(self, lib: RFExcelLibrary, tmp_path):
        path = str(tmp_path / "notes.txt")
        with pytest.raises(FileFormatNotSupportedException):
            lib.create_workbook(path)

    def test_create_ods_raises_format_not_supported(self, lib: RFExcelLibrary, tmp_path):
        path = str(tmp_path / "sheet.ods")
        with pytest.raises(FileFormatNotSupportedException):
            lib.create_workbook(path)

    def test_active_workbook_unchanged_after_failed_create(self, lib: RFExcelLibrary, tmp_path):
        """A failed create must not overwrite a currently active workbook."""
        lib.load_workbook(XLSX_FILE)
        active_before = lib._active_workbook
        with pytest.raises(FileFormatNotSupportedException):
            lib.create_workbook(str(tmp_path / "bad.txt"))
        assert lib._active_workbook is active_before


# ─── edge cases ───────────────────────────────────────────────────────────────

class TestCreateWorkbookEdge:

    def test_create_xlsx_with_nested_new_directories(self, lib: RFExcelLibrary, tmp_path):
        """Parent directories that do not exist must be created automatically."""
        path = str(tmp_path / "a" / "b" / "c" / "deep.xlsx")
        lib.create_workbook(path)
        assert Path(path).exists()

    def test_created_xlsx_can_be_loaded_afterwards(self, lib: RFExcelLibrary, tmp_path):
        path = str(tmp_path / "roundtrip.xlsx")
        lib.create_workbook(path)
        lib.close()
        lib.load_workbook(path)
        assert lib._active_workbook is not None

    def test_created_csv_can_be_loaded_afterwards(self, lib: RFExcelLibrary, tmp_path):
        path = str(tmp_path / "roundtrip.csv")
        lib.create_workbook(path)
        lib.close()
        lib.load_workbook(path)
        assert lib._active_workbook is not None

    def test_two_different_workbooks_created_independently(self, lib: RFExcelLibrary, tmp_path):
        path_a = str(tmp_path / "a.xlsx")
        path_b = str(tmp_path / "b.xlsx")
        lib.create_workbook(path_a)
        lib.close()
        lib.create_workbook(path_b)
        assert Path(path_a).exists()
        assert Path(path_b).exists()
