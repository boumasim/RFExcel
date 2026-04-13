from typing import Any, cast

import pytest

from rfexcel.exception.library_exceptions import (
	StreamingViolationException,
	WorkbookNotOpenException,
)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import (
	BACKEND_NAMES,
	SHEET1_HEADER_MAP_DICT,
	SHEET1_HEADERS,
	SHEET1_ROWS,
	STREAMING_BACKENDS,
	XLSX_EDIT,
	open_backend,
)

SHEET1_ROWS_AS_LISTS = [
    ["P-200", "Wireless Mouse", 25.5, "Warehouse A, Shelf 2"],
    ["P-201", "Keyboard, Mechanical", 89.99, "Store Front"],
    ["P-202", "Monitor 24-inch", 150, "Paris, France"],
    ["P-203", "USB Cable", 5.99, "OnlineP"]
]


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_get_row_without_headers_returns_expected_list_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    result = lib.get_row(2)
    assert isinstance(result, list)
    assert result == SHEET1_ROWS_AS_LISTS[0]


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_get_row_with_headers_returns_expected_dict_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    headers = SHEET1_HEADERS
    result = lib.get_row(2, headers=headers)
    assert isinstance(result, dict)
    assert result == SHEET1_ROWS[0]


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_get_row_out_of_bounds_returns_empty_list_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    assert lib.get_row(9999) == []


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_repeated_get_row_call_matches_backend_mode_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    first_result = lib.get_row(2)

    if backend_name in STREAMING_BACKENDS:
        with pytest.raises(StreamingViolationException):
            lib.get_row(2)
        return

    assert lib.get_row(2) == first_result


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_reading_earlier_row_after_forward_reads_matches_backend_mode_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    first_row = lib.get_row(1)
    second_row = lib.get_row(2)

    if backend_name in STREAMING_BACKENDS:
        with pytest.raises(StreamingViolationException):
            lib.get_row(1)
        return

    assert lib.get_row(1) == first_row
    assert lib.get_row(2) == second_row


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_get_row_with_partial_headers_maps_only_requested_columns_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    requested_headers = SHEET1_HEADERS[:2]

    result = lib.get_row(2, headers=requested_headers)
    typed_row = cast(dict[str, Any], result)

    assert isinstance(result, dict)
    assert list(typed_row.keys()) == requested_headers
    assert len(typed_row) == 2


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_get_row_with_empty_headers_returns_list_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    result = lib.get_row(2, headers=[])
    assert isinstance(result, list)


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_row_zero_returns_empty_list_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    assert lib.get_row(0) == []


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_get_row_with_headers_uses_expected_header_order_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    headers = SHEET1_HEADER_MAP_DICT

    result = lib.get_row(2, headers=headers)

    assert result == SHEET1_ROWS[0]


def test_get_row_raises_when_no_workbook_is_loaded(lib: RFExcelLibrary) -> None:
    with pytest.raises(WorkbookNotOpenException):
        lib.get_row(1)


def test_get_row_raises_after_close(lib: RFExcelLibrary) -> None:
    open_backend(lib, XLSX_EDIT)
    lib.close()
    with pytest.raises(WorkbookNotOpenException):
        lib.get_row(1)
