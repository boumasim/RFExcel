from pathlib import Path

import pytest

from rfexcel.exception.library_exceptions import (
	HeadersNotDeterminedException,
	NullComponentException,
	StreamingViolationException,
	WorkbookNotOpenException,
)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.integration.data.add_data import (
	ADD_ROWS_PARTIAL_INPUT,
	EXPECTED_ADD_ROWS_PARTIAL,
	EXPECTED_ORDERED_ROWS,
	ORDERED_ROWS,
)
from tests.pyth.test_data import BACKEND_NAMES, BACKENDS, STREAMING_BACKENDS, load_backend_copy


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_append_rows_matches_backend_mode_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)
    rows_input = list(ORDERED_ROWS)
    expected_first_row, expected_second_row = EXPECTED_ORDERED_ROWS

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.append_rows(rows_input)
        return

    rows_before = lib.get_rows()
    lib.append_rows(rows_input)
    rows_after = lib.get_rows()

    assert len(rows_after) == len(rows_before) + 2
    assert rows_after[-2] == expected_first_row
    assert rows_after[-1] == expected_second_row


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_append_rows_empty_list_is_noop_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)
    rows_before = lib.get_rows()
    lib.append_rows([])
    if backend_name in STREAMING_BACKENDS:
        with pytest.raises(StreamingViolationException):
            lib.get_rows()
        assert len(rows_before) > 0
        return
    assert lib.get_rows() == rows_before


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_append_rows_partial_rows_fill_missing_values_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)
    rows_input = ADD_ROWS_PARTIAL_INPUT
    expected_first_row, expected_second_row = EXPECTED_ADD_ROWS_PARTIAL

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.append_rows(rows_input)
        return

    lib.append_rows(rows_input)
    rows_after = lib.get_rows()
    assert rows_after[-2] == expected_first_row
    assert rows_after[-1] == expected_second_row


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_append_rows_header_row_out_of_range_raises_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)
    with pytest.raises(HeadersNotDeterminedException):
        lib.append_rows(list(ORDERED_ROWS), header_row=9999)

def test_append_rows_raises_when_no_workbook_is_loaded(lib: RFExcelLibrary) -> None:
    with pytest.raises(WorkbookNotOpenException):
        lib.append_rows([])
