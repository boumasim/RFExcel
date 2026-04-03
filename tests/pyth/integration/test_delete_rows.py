from pathlib import Path
from typing import Any, cast

import pytest

from rfexcel.exception.library_exceptions import (
    HeadersNotDeterminedException, NullComponentException,
    WorkbookNotOpenException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.integration.data.delete_data import (
    DELETE_ROWS_DUPLICATE_MATCH_CRITERIA,
    DELETE_ROWS_DUPLICATE_UPDATE_CRITERIA, DELETE_ROWS_DUPLICATE_UPDATE_VALUES,
    DELETE_ROWS_NO_MATCH_CRITERIA, DELETE_ROWS_NUMERIC_MATCH_CRITERIA,
    DELETE_ROWS_PARTIAL_MATCH_CRITERIA, DELETE_ROWS_SECOND_MATCH_CRITERIA,
    DELETE_ROWS_SINGLE_MATCH_CRITERIA, EXPECTED_DUPLICATE_DELETE_COUNT,
    EXPECTED_DUPLICATE_ROWS_REMAINING_AFTER_ONE_ROW_DELETE,
    EXPECTED_PARTIAL_MATCH_DELETE_COUNT,
    EXPECTED_ROWS_AFTER_DELETE_NUMERIC_MATCH,
    EXPECTED_ROWS_AFTER_DELETE_SINGLE_MATCH)
from tests.pyth.test_data import (BACKEND_NAMES, BACKENDS, load_backend_copy,
                                  open_backend)


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_delete_rows_matches_backend_mode_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.delete_rows(search_criteria=DELETE_ROWS_SINGLE_MATCH_CRITERIA)
        return

    rows_before = lib.get_rows()
    count = lib.delete_rows(search_criteria=DELETE_ROWS_SINGLE_MATCH_CRITERIA)
    rows_after = lib.get_rows()

    assert count == 1
    assert len(rows_after) == len(rows_before) - count
    assert rows_after == EXPECTED_ROWS_AFTER_DELETE_SINGLE_MATCH


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_delete_rows_no_match_keeps_data_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    if BACKENDS[backend_name][1]:
        assert lib.delete_rows(search_criteria=DELETE_ROWS_NO_MATCH_CRITERIA) == 0
        return

    rows_before = lib.get_rows()
    count = lib.delete_rows(search_criteria=DELETE_ROWS_NO_MATCH_CRITERIA)

    assert count == 0
    assert lib.get_rows() == rows_before


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_delete_rows_deletes_all_matching_rows_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    if BACKENDS[backend_name][1]:
        assert lib.delete_rows(search_criteria=DELETE_ROWS_DUPLICATE_MATCH_CRITERIA) == 0
        return

    for criteria in DELETE_ROWS_DUPLICATE_UPDATE_CRITERIA:
        lib.update_values(
            search_criteria=criteria,
            values=DELETE_ROWS_DUPLICATE_UPDATE_VALUES,
        )

    count = lib.delete_rows(search_criteria=DELETE_ROWS_DUPLICATE_MATCH_CRITERIA)
    rows_after = cast(list[dict[str, Any]], lib.get_rows())

    assert count == EXPECTED_DUPLICATE_DELETE_COUNT
    assert all(row["Location"] != DELETE_ROWS_DUPLICATE_UPDATE_VALUES["Location"] for row in rows_after)


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_delete_rows_one_row_deletes_single_match_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    if BACKENDS[backend_name][1]:
        assert lib.delete_rows(search_criteria=DELETE_ROWS_DUPLICATE_MATCH_CRITERIA, one_row=True) == 0
        return

    for criteria in DELETE_ROWS_DUPLICATE_UPDATE_CRITERIA:
        lib.update_values(
            search_criteria=criteria,
            values=DELETE_ROWS_DUPLICATE_UPDATE_VALUES,
        )

    count = lib.delete_rows(search_criteria=DELETE_ROWS_DUPLICATE_MATCH_CRITERIA, one_row=True)
    rows_after = cast(list[dict[str, Any]], lib.get_rows())
    duplicate_value = DELETE_ROWS_DUPLICATE_UPDATE_VALUES["Location"]

    assert count == 1
    assert sum(1 for row in rows_after if row["Location"] == duplicate_value) == (
        EXPECTED_DUPLICATE_ROWS_REMAINING_AFTER_ONE_ROW_DELETE
    )


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_delete_rows_partial_match_behaves_consistently_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.delete_rows(search_criteria=DELETE_ROWS_PARTIAL_MATCH_CRITERIA, partial_match=True)
        return

    rows_before = cast(list[dict[str, Any]], lib.get_rows())
    count = lib.delete_rows(
        search_criteria=DELETE_ROWS_PARTIAL_MATCH_CRITERIA,
        partial_match=True,
    )
    rows_after = cast(list[dict[str, Any]], lib.get_rows())

    assert count == EXPECTED_PARTIAL_MATCH_DELETE_COUNT
    assert len(rows_after) == len(rows_before) - count
    assert all("Warehouse" not in cast(str, row["Location"]) for row in rows_after)


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_delete_rows_numeric_string_search_matches_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.delete_rows(search_criteria=DELETE_ROWS_NUMERIC_MATCH_CRITERIA)
        return

    count = lib.delete_rows(search_criteria=DELETE_ROWS_NUMERIC_MATCH_CRITERIA)

    assert count == 1
    assert lib.get_rows() == EXPECTED_ROWS_AFTER_DELETE_NUMERIC_MATCH


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_delete_rows_second_single_match_decreases_count_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    if BACKENDS[backend_name][1]:
        with pytest.raises(NullComponentException):
            lib.delete_rows(search_criteria=DELETE_ROWS_SECOND_MATCH_CRITERIA)
        return

    rows_before = lib.get_rows()
    count = lib.delete_rows(search_criteria=DELETE_ROWS_SECOND_MATCH_CRITERIA)

    assert count == 1
    assert len(lib.get_rows()) == len(rows_before) - 1


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_delete_rows_header_row_out_of_range_raises_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)

    with pytest.raises(HeadersNotDeterminedException):
        lib.delete_rows(
            search_criteria=DELETE_ROWS_SINGLE_MATCH_CRITERIA,
            header_row=9999,
        )


def test_delete_rows_raises_when_no_workbook_is_loaded(lib: RFExcelLibrary) -> None:
    with pytest.raises(WorkbookNotOpenException):
        lib.delete_rows(search_criteria=DELETE_ROWS_SINGLE_MATCH_CRITERIA)
