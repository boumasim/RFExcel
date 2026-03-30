from pathlib import Path
from typing import Any, cast

import pytest

from rfexcel.exception.library_exceptions import (FileSaveException,
                                                  NotSupportedInReadOnlyMode,
                                                  WorkbookNotOpenException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import (BACKEND_NAMES, BACKENDS, CSV_EDIT, XLS_EDIT,
                                  load_backend_copy, EDITABLE_BACKENDS, XLSX_FORMAT)

SAVE_AS_SUFFIX_BY_BACKEND: dict[str, str] = {
    backend_name: (XLSX_FORMAT if backend_name == XLS_EDIT else Path(BACKENDS[backend_name][0]).suffix.lstrip("."))
    for backend_name in BACKEND_NAMES
}

COPY_ONLY_MARKER = "OnlyInCopy"
SECOND_SAVE_MARKER = "SecondSave"


def build_csv_marker_row(marker: str) -> dict[str, Any]:
    return {
        "Product ID": f"P-{marker}",
        "Description": marker,
        "Price": 1.23,
        "Location": marker,
    }


def read_rows(path: str) -> list[dict[str, Any]]:
    reloaded_library = RFExcelLibrary()
    reloaded_library.load_workbook(path)
    try:
        return cast(list[dict[str, Any]], reloaded_library.get_rows())
    finally:
        reloaded_library.close()


def read_sheet_names(path: str) -> list[str]:
    reloaded_library = RFExcelLibrary()
    reloaded_library.load_workbook(path)
    try:
        return reloaded_library.list_sheet_names()
    finally:
        reloaded_library.close()


def apply_marker_mutation(lib: RFExcelLibrary, backend_name: str, marker: str) -> None:
    if backend_name == CSV_EDIT:
        lib.append_row(build_csv_marker_row(marker))
        return

    lib.add_sheet(marker)


def assert_marker_persisted(path: str, backend_name: str, marker: str) -> None:
    if backend_name == CSV_EDIT:
        assert read_rows(path)[-1] == build_csv_marker_row(marker)
        return

    assert marker in read_sheet_names(path)


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_save_workbook_matches_backend_capabilities_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    loaded_path = load_backend_copy(lib, backend_name, tmp_path)

    if BACKENDS[backend_name][1]:
        with pytest.raises(NotSupportedInReadOnlyMode):
            lib.save_workbook()
        return

    reload_path = loaded_path

    if backend_name == CSV_EDIT:
        expected_rows = cast(list[dict[str, Any]], lib.get_rows())
        lib.save_workbook()
        lib.close()
        assert read_rows(reload_path) == expected_rows
        return

    apply_marker_mutation(lib, backend_name, COPY_ONLY_MARKER)

    if backend_name == XLS_EDIT:
        reload_path = str(tmp_path / "persisted.xlsx")
        lib.save_workbook(reload_path)
    else:
        lib.save_workbook()

    lib.close()
    assert COPY_ONLY_MARKER in read_sheet_names(reload_path)


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_save_as_matches_backend_capabilities_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)
    new_path = tmp_path / f"copy.{SAVE_AS_SUFFIX_BY_BACKEND[backend_name]}"

    if BACKENDS[backend_name][1]:
        with pytest.raises(NotSupportedInReadOnlyMode):
            lib.save_workbook(str(new_path))
        return

    lib.save_workbook(str(new_path))
    assert new_path.exists()


@pytest.mark.parametrize("backend_name", EDITABLE_BACKENDS, ids=EDITABLE_BACKENDS)
def test_save_as_keeps_original_copy_unchanged_for_editable_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    loaded_path = load_backend_copy(lib, backend_name, tmp_path)
    new_path = tmp_path / f"isolated.{SAVE_AS_SUFFIX_BY_BACKEND[backend_name]}"

    if backend_name == CSV_EDIT:
        rows_before = cast(list[dict[str, Any]], lib.get_rows())
        apply_marker_mutation(lib, backend_name, COPY_ONLY_MARKER)
        lib.save_workbook(str(new_path))
        lib.close()

        assert read_rows(loaded_path) == rows_before
        assert_marker_persisted(str(new_path), backend_name, COPY_ONLY_MARKER)
        return

    sheet_names_before = lib.list_sheet_names()
    apply_marker_mutation(lib, backend_name, COPY_ONLY_MARKER)
    lib.save_workbook(str(new_path))
    lib.close()

    assert read_sheet_names(loaded_path) == sheet_names_before
    assert_marker_persisted(str(new_path), backend_name, COPY_ONLY_MARKER)


@pytest.mark.parametrize("backend_name", EDITABLE_BACKENDS, ids=EDITABLE_BACKENDS)
def test_save_as_updates_active_path_for_subsequent_saves(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)
    new_path = tmp_path / f"moved.{SAVE_AS_SUFFIX_BY_BACKEND[backend_name]}"

    lib.save_workbook(str(new_path))
    apply_marker_mutation(lib, backend_name, SECOND_SAVE_MARKER)
    lib.save_workbook()
    lib.close()

    assert_marker_persisted(str(new_path), backend_name, SECOND_SAVE_MARKER)


@pytest.mark.parametrize("backend_name", EDITABLE_BACKENDS, ids=EDITABLE_BACKENDS)
def test_save_to_nonexistent_dir_raises_for_editable_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)
    bad_path = tmp_path / "no_such_dir" / f"out.{SAVE_AS_SUFFIX_BY_BACKEND[backend_name]}"

    with pytest.raises(FileSaveException):
        lib.save_workbook(str(bad_path))


def test_save_raises_when_no_workbook_is_loaded(lib: RFExcelLibrary) -> None:
    with pytest.raises(WorkbookNotOpenException):
        lib.save_workbook()
