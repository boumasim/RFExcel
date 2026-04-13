from typing import Any, cast

import pytest

from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import (
	SHEET1_HEADERS,
	SHEET1_ROWS,
	SHEET3_NAME,
	SHIFTED_ROW_START_IDX,
	XLS_EDIT,
	XLS_ON_DEMAND,
	XLSX_EDIT,
	XLSX_STREAM,
	open_backend,
)

SPARSE_BACKENDS = [XLSX_EDIT, XLSX_STREAM, XLS_EDIT, XLS_ON_DEMAND]

@pytest.mark.parametrize("backend_name", SPARSE_BACKENDS, ids=SPARSE_BACKENDS)
def test_sparse_sheet_returns_expected_rows_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)

    lib.switch_sheet(SHEET3_NAME)
    assert lib.get_rows(header_row=SHIFTED_ROW_START_IDX) == SHEET1_ROWS


@pytest.mark.parametrize("backend_name", SPARSE_BACKENDS, ids=SPARSE_BACKENDS)
def test_sparse_sheet_has_expected_row_count_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)

    lib.switch_sheet(SHEET3_NAME)
    assert len(lib.get_rows(header_row=SHIFTED_ROW_START_IDX)) == len(SHEET1_ROWS)


@pytest.mark.parametrize("backend_name", SPARSE_BACKENDS, ids=SPARSE_BACKENDS)
def test_sparse_sheet_headers_match_sheet1_headers_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)

    lib.switch_sheet(SHEET3_NAME)
    first_row = cast(dict[str, Any], lib.get_rows(header_row=SHIFTED_ROW_START_IDX)[0])
    assert list(first_row.keys()) == SHEET1_HEADERS
