from pathlib import Path

import pytest

from rfexcel.exception.library_exceptions import (
	NullComponentException,
	RowIndexOutOfBoundsException,
	WorkbookNotOpenException,
)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.integration.data.delete_data import DELETE_ROW_SCENARIOS
from tests.pyth.test_data import BACKEND_NAMES, BACKENDS, load_backend_copy


@pytest.mark.parametrize(
    ("scenario_name", "row_index", "expected_rows"),
    DELETE_ROW_SCENARIOS,
    ids=[scenario_name for scenario_name, _, _ in DELETE_ROW_SCENARIOS],
)
@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_delete_row_matches_backend_mode_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    scenario_name: str,
    row_index: int,
    expected_rows: list[dict[str, object]],
    tmp_path: Path,
) -> None:
    del scenario_name
    load_backend_copy(lib, backend_name, tmp_path)

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.delete_row(row_index)
        return

    rows_before = lib.get_rows()
    lib.delete_row(row_index)
    rows_after = lib.get_rows()

    assert len(rows_after) == len(rows_before) - 1
    assert rows_after == expected_rows


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_delete_header_row_matches_backend_mode_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.delete_row(1)
        return

    expected_first_row = lib.get_row(2)
    lib.delete_row(1)
    assert lib.get_row(1) == expected_first_row


@pytest.mark.parametrize("row_index", (0, -1), ids=["zero", "negative_one"])
@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_delete_row_invalid_indices_match_backend_mode_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    row_index: int,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    with pytest.raises(RowIndexOutOfBoundsException):
        lib.delete_row(row_index)


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_delete_row_out_of_bounds_matches_backend_mode_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    with pytest.raises(RowIndexOutOfBoundsException):
        lib.delete_row(9999)


def test_delete_row_raises_when_no_workbook_is_loaded(lib: RFExcelLibrary) -> None:
    with pytest.raises(WorkbookNotOpenException):
        lib.delete_row(2)
