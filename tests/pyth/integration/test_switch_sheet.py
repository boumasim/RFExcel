from typing import Any, cast

import pytest

from rfexcel.exception.library_exceptions import (
    OperationNotSupportedForFormat, SheetDoesNotExistException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import (BACKEND_NAMES, CSV_EDIT, CSV_STREAM, SHEET1_ROWS, SHEET2_EXPECTED_ROW_COUNT, SHEET2_HEADERS, SHEET2_NAME, SHEET2_ROWS, open_backend, SHEET1_NAME)

CSV_BACKENDS = [CSV_EDIT, CSV_STREAM]

PRIMARY_FIRST_ROW: dict[str, Any] = SHEET1_ROWS[0]

SECONDARY_FIRST_ROW: dict[str, Any] = SHEET2_ROWS[0]



@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_switch_to_secondary_sheet_returns_expected_first_row_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)

    if backend_name in CSV_BACKENDS:
        with pytest.raises(OperationNotSupportedForFormat):
            lib.switch_sheet("anything")
        return

    lib.switch_sheet(SHEET2_NAME)
    assert lib.get_rows()[0] == SECONDARY_FIRST_ROW


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_switch_to_secondary_sheet_returns_expected_row_count_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)

    if backend_name in CSV_BACKENDS:
        with pytest.raises(OperationNotSupportedForFormat):
            lib.switch_sheet("anything")
        return

    lib.switch_sheet(SHEET2_NAME)
    assert len(lib.get_rows()) == SHEET2_EXPECTED_ROW_COUNT


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_switch_back_to_primary_sheet_restores_first_row_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)

    if backend_name in CSV_BACKENDS:
        with pytest.raises(OperationNotSupportedForFormat):
            lib.switch_sheet("anything")
        return

    lib.switch_sheet(SHEET2_NAME)
    lib.switch_sheet(SHEET1_NAME)
    assert lib.get_rows()[0] == PRIMARY_FIRST_ROW


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_switch_sheet_updates_headers_consistently_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)

    if backend_name in CSV_BACKENDS:
        with pytest.raises(OperationNotSupportedForFormat):
            lib.switch_sheet("anything")
        return

    lib.switch_sheet(SHEET2_NAME)
    first_row = cast(dict[str, Any], lib.get_rows()[0])
    assert list(first_row.keys()) == SHEET2_HEADERS


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_switch_sheet_after_read_resets_reader_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)

    if backend_name in CSV_BACKENDS:
        lib.get_rows()
        with pytest.raises(OperationNotSupportedForFormat):
            lib.switch_sheet("anything")
        return

    lib.get_rows()
    lib.switch_sheet(SHEET2_NAME)
    assert lib.get_rows()[0] == SECONDARY_FIRST_ROW


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_switch_to_nonexistent_sheet_raises_expected_exception_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)

    if backend_name in CSV_BACKENDS:
        with pytest.raises(OperationNotSupportedForFormat):
            lib.switch_sheet("DoesNotExist")
        return

    with pytest.raises(SheetDoesNotExistException, match="DoesNotExist"):
        lib.switch_sheet("DoesNotExist")


@pytest.mark.parametrize("backend_name", [CSV_EDIT], ids=[CSV_EDIT])
def test_csv_switch_sheet_error_mentions_csv(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    with pytest.raises(OperationNotSupportedForFormat, match="(?i)csv"):
        lib.switch_sheet("anything")
