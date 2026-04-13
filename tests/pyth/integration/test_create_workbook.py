from pathlib import Path

import pytest

from rfexcel.exception.library_exceptions import (
	FileAlreadyExistsException,
	FileFormatNotSupportedException,
	HeadersNotDeterminedException,
)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import CSV_FORMAT, FORMAT_FILE, FORMAT_LIST, XLS_FORMAT, XLSX_FORMAT

BACKEND_CREATABLE_BY_FORMAT: dict[str, bool] = {
    XLSX_FORMAT: True,
    XLS_FORMAT:  False,
    CSV_FORMAT:  True,
}

CREATABLE_FORMATS     = [fmt for fmt in FORMAT_LIST if BACKEND_CREATABLE_BY_FORMAT[fmt]]
NON_CREATABLE_FORMATS = [fmt for fmt in FORMAT_LIST if not BACKEND_CREATABLE_BY_FORMAT[fmt]]

@pytest.mark.parametrize("fmt", CREATABLE_FORMATS, ids=CREATABLE_FORMATS)
def test_create_sets_active_workbook(
    lib: RFExcelLibrary, tmp_path: Path, fmt: str
) -> None:
    lib.create_workbook(str(tmp_path / f"new.{fmt}"))
    assert lib.active_workbook is not None


@pytest.mark.parametrize("fmt", CREATABLE_FORMATS, ids=CREATABLE_FORMATS)
def test_create_produces_file_on_disk(
    lib: RFExcelLibrary, tmp_path: Path, fmt: str
) -> None:
    path = tmp_path / f"created.{fmt}"
    lib.create_workbook(str(path))
    assert path.exists()


@pytest.mark.parametrize("fmt", CREATABLE_FORMATS, ids=CREATABLE_FORMATS)
def test_created_empty_file_raises_on_get_rows(
    lib: RFExcelLibrary, tmp_path: Path, fmt: str
) -> None:
    lib.create_workbook(str(tmp_path / f"empty.{fmt}"))
    with pytest.raises(HeadersNotDeterminedException):
        lib.get_rows()


@pytest.mark.parametrize("fmt", CREATABLE_FORMATS, ids=CREATABLE_FORMATS)
def test_create_on_existing_file_raises(
    lib: RFExcelLibrary, tmp_path: Path, fmt: str
) -> None:
    path = str(tmp_path / f"existing.{fmt}")
    lib.create_workbook(path)
    lib.close()
    with pytest.raises(FileAlreadyExistsException):
        lib.create_workbook(path)


@pytest.mark.parametrize("fmt", NON_CREATABLE_FORMATS, ids=NON_CREATABLE_FORMATS)
def test_create_non_creatable_format_raises(
    lib: RFExcelLibrary, tmp_path: Path, fmt: str
) -> None:
    with pytest.raises(FileFormatNotSupportedException):
        lib.create_workbook(str(tmp_path / f"file.{fmt}"))


@pytest.mark.parametrize("fmt", ["txt", "ods"], ids=["txt", "ods"])
def test_create_unsupported_format_raises(
    lib: RFExcelLibrary, tmp_path: Path, fmt: str
) -> None:
    with pytest.raises(FileFormatNotSupportedException):
        lib.create_workbook(str(tmp_path / f"file.{fmt}"))


@pytest.mark.parametrize("fmt", FORMAT_LIST, ids=FORMAT_LIST)
def test_failed_create_clears_previously_active_workbook(
    lib: RFExcelLibrary, tmp_path: Path, fmt: str
) -> None:
    lib.load_workbook(FORMAT_FILE[fmt])
    with pytest.raises(FileFormatNotSupportedException):
        lib.create_workbook(str(tmp_path / "bad.ods"))
    assert lib.active_workbook is None


@pytest.mark.parametrize("fmt", CREATABLE_FORMATS, ids=CREATABLE_FORMATS)
def test_create_with_nested_new_directories(
    lib: RFExcelLibrary, tmp_path: Path, fmt: str
) -> None:
    path = tmp_path / "a" / "b" / "c" / f"deep.{fmt}"
    lib.create_workbook(str(path))
    assert path.exists()


@pytest.mark.parametrize("fmt", CREATABLE_FORMATS, ids=CREATABLE_FORMATS)
def test_created_file_can_be_loaded_afterwards(
    lib: RFExcelLibrary, tmp_path: Path, fmt: str
) -> None:
    path = str(tmp_path / f"roundtrip.{fmt}")
    lib.create_workbook(path)
    lib.load_workbook(path)
    assert lib.active_workbook is not None


@pytest.mark.parametrize("fmt", CREATABLE_FORMATS, ids=CREATABLE_FORMATS)
def test_two_workbooks_created_independently(
    lib: RFExcelLibrary, tmp_path: Path, fmt: str
) -> None:
    path_a = tmp_path / f"a.{fmt}"
    path_b = tmp_path / f"b.{fmt}"
    lib.create_workbook(str(path_a))
    lib.create_workbook(str(path_b))
    assert path_a.exists()
    assert path_b.exists()
