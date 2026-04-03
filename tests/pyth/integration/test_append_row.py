from pathlib import Path

import pytest

from rfexcel.exception.library_exceptions import (
    HeadersNotDeterminedException, NullComponentException,
    WorkbookNotOpenException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.integration.data.add_data import (
    EXPECTED_FULL_ROW, EXPECTED_ORDERED_ROWS,
    EXPECTED_PARTIAL_ROW, EXPECTED_UNKNOWN_KEY_ROW,
    FULL_ROW, ORDERED_ROWS, PARTIAL_ROW,
    UNKNOWN_KEY_ROW)
from tests.pyth.test_data import BACKEND_NAMES, BACKENDS, XLS_EDIT, load_backend_copy


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_append_row_matches_backend_mode_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)
    row_data = FULL_ROW

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.append_row(row_data)
        return

    rows_before = lib.get_rows()
    lib.append_row(row_data)
    rows_after = lib.get_rows()

    assert len(rows_after) == len(rows_before) + 1
    assert rows_after[-1] == EXPECTED_FULL_ROW


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_partial_append_row_matches_backend_mode_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)
    row_data = PARTIAL_ROW

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.append_row(row_data)
        return

    lib.append_row(row_data)
    assert lib.get_rows()[-1] == EXPECTED_PARTIAL_ROW


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_unknown_keys_are_ignored_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)
    row_data = UNKNOWN_KEY_ROW

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.append_row(row_data)
        return

    rows_before = lib.get_rows()
    lib.append_row(row_data)
    rows_after = lib.get_rows()

    assert len(rows_after) == len(rows_before) + 1
    assert rows_after[-1] == EXPECTED_UNKNOWN_KEY_ROW


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_appended_row_is_persisted_after_save_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    loaded_path = load_backend_copy(lib, backend_name, tmp_path)

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.append_row(FULL_ROW)
        return

    lib.append_row(FULL_ROW)

    reload_path = loaded_path
    if backend_name == XLS_EDIT:
        reload_path = str(tmp_path / "result.xlsx")
        lib.save_workbook(reload_path)
    else:
        lib.save_workbook()
    lib.close()

    reloaded_library = RFExcelLibrary()
    reloaded_library.load_workbook(reload_path)
    try:
        assert reloaded_library.get_rows()[-1] == EXPECTED_FULL_ROW
    finally:
        reloaded_library.close()


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_multiple_appended_rows_preserve_order_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)
    first_row, second_row = ORDERED_ROWS
    expected_first_row, expected_second_row = EXPECTED_ORDERED_ROWS

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.append_row(first_row)
        return

    lib.append_row(first_row)
    lib.append_row(second_row)
    rows = lib.get_rows()

    assert rows[-2] == expected_first_row
    assert rows[-1] == expected_second_row


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_header_row_out_of_range_raises_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)
    with pytest.raises(HeadersNotDeterminedException):
        lib.append_row(FULL_ROW, header_row=9999)


def test_append_row_raises_when_no_workbook_is_loaded(lib: RFExcelLibrary) -> None:
    with pytest.raises(WorkbookNotOpenException):
        lib.append_row(FULL_ROW)