from pathlib import Path
from typing import cast

import pytest

from rfexcel.exception.library_exceptions import (
	HeadersNotDeterminedException,
	NullComponentException,
	RowIndexOutOfBoundsException,
	WorkbookNotOpenException,
)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.integration.data.add_data import (
	EXPECTED_FULL_ROW,
	EXPECTED_PARTIAL_ROW,
	EXPECTED_UNKNOWN_KEY_ROW,
	FULL_ROW,
	PARTIAL_ROW,
	UNKNOWN_KEY_ROW,
)
from tests.pyth.test_data import (
	BACKEND_NAMES,
	BACKENDS,
	SHEET1_ROWS,
	SHEET3_NAME,
	SHIFTED_ROW_START_IDX,
	XLS_EDIT,
	XLSX_EDIT,
	load_backend_copy,
)


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_insert_row_matches_backend_mode_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.insert_row(FULL_ROW, row=2)
        return

    rows_before = lib.get_rows()
    first_row_before = cast(str, rows_before[0]["Product ID"])

    lib.insert_row(FULL_ROW, row=2)
    rows_after = lib.get_rows()

    assert len(rows_after) == len(rows_before) + 1
    assert rows_after[0] == EXPECTED_FULL_ROW
    assert rows_after[1]["Product ID"] == first_row_before


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_partial_insert_row_matches_backend_mode_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.insert_row(PARTIAL_ROW, row=2)
        return

    lib.insert_row(PARTIAL_ROW, row=2)
    assert lib.get_rows()[0] == EXPECTED_PARTIAL_ROW


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_insert_row_ignores_unknown_keys_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.insert_row(UNKNOWN_KEY_ROW, row=2)
        return

    rows_before = lib.get_rows()
    lib.insert_row(UNKNOWN_KEY_ROW, row=2)
    rows_after = lib.get_rows()

    assert len(rows_after) == len(rows_before) + 1
    assert rows_after[0] == EXPECTED_UNKNOWN_KEY_ROW


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_inserted_row_is_persisted_after_save_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    loaded_path = load_backend_copy(lib, backend_name, tmp_path)

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.insert_row(FULL_ROW, row=2)
        return

    lib.insert_row(FULL_ROW, row=2)

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
        assert reloaded_library.get_rows()[0] == EXPECTED_FULL_ROW
    finally:
        reloaded_library.close()


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_insert_row_at_last_data_position_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.insert_row(PARTIAL_ROW, row=2)
        return

    rows_before = lib.get_rows()
    last_row_before = cast(str, rows_before[-1]["Product ID"])
    insert_at = len(rows_before) + 1

    lib.insert_row({"Product ID": "P-LAST"}, row=insert_at)
    rows_after = lib.get_rows()

    assert rows_after[-2]["Product ID"] == "P-LAST"
    assert rows_after[-1]["Product ID"] == last_row_before


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_insert_row_with_invalid_header_row_raises_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    with pytest.raises(HeadersNotDeterminedException):
        lib.insert_row(FULL_ROW, row=9999, header_row=9998)


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_insert_row_equal_to_header_row_raises_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    with pytest.raises(RowIndexOutOfBoundsException):
        lib.insert_row(FULL_ROW, row=1, header_row=1)


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_insert_row_less_than_header_row_raises_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    with pytest.raises(RowIndexOutOfBoundsException):
        lib.insert_row(FULL_ROW, row=1, header_row=2)


def test_insert_row_raises_when_no_workbook_is_loaded(lib: RFExcelLibrary) -> None:
    with pytest.raises(WorkbookNotOpenException):
        lib.insert_row(FULL_ROW, row=2)


def test_insert_row_on_xls_keeps_original_copy_unchanged(
    lib: RFExcelLibrary,
    tmp_path: Path,
) -> None:
    xls_path = load_backend_copy(lib, XLS_EDIT, tmp_path)
    rows_before = lib.get_rows()
    lib.close()

    lib.load_workbook(xls_path)
    lib.insert_row(FULL_ROW, row=2)
    lib.save_workbook(str(tmp_path / "out.xlsx"))
    lib.close()

    verifier = RFExcelLibrary()
    verifier.load_workbook(xls_path)
    try:
        assert verifier.get_rows() == rows_before
    finally:
        verifier.close()


@pytest.mark.parametrize("backend_name", [XLSX_EDIT, XLS_EDIT], ids=[XLSX_EDIT, XLS_EDIT])
def test_insert_row_on_sheet3_shifted_data_for_xls_and_xlsx(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)
    lib.switch_sheet(SHEET3_NAME)

    rows_before = lib.get_rows(header_row=SHIFTED_ROW_START_IDX)
    assert rows_before == SHEET1_ROWS

    insert_at = SHIFTED_ROW_START_IDX + 1
    lib.insert_row(FULL_ROW, row=insert_at, header_row=SHIFTED_ROW_START_IDX)

    rows_after = lib.get_rows(header_row=SHIFTED_ROW_START_IDX)

    assert len(rows_after) == len(rows_before) + 1
    assert rows_after[0] == EXPECTED_FULL_ROW
    assert rows_after[1]["Product ID"] == rows_before[0]["Product ID"]