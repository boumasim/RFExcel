from pathlib import Path
from typing import Any, cast

import pytest

from rfexcel.exception.library_exceptions import (
	HeadersNotDeterminedException,
	NullComponentException,
	WorkbookNotOpenException,
)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import (
	BACKEND_NAMES,
	EDITABLE_BACKENDS,
	SHEET1_ROWS,
	STREAMING_BACKENDS,
	load_backend_copy,
	open_backend,
)


@pytest.mark.parametrize("backend_name", EDITABLE_BACKENDS, ids=EDITABLE_BACKENDS)
def test_update_values_updates_single_matching_row_for_editable_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    count = lib.update_values(
        search_criteria={"Product ID": "P-200"},
        values={"Location": "Updated Location"},
    )

    assert count == 1
    rows = cast(list[dict[str, Any]], lib.get_rows())
    row = next(r for r in rows if r["Product ID"] == "P-200")
    assert row["Location"] == "Updated Location"


@pytest.mark.parametrize("backend_name", EDITABLE_BACKENDS, ids=EDITABLE_BACKENDS)
def test_update_values_leaves_unspecified_columns_unchanged_for_editable_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    count = lib.update_values(
        search_criteria={"Product ID": "P-201"},
        values={"Location": "Only Location Changed"},
    )

    assert count == 1
    rows = cast(list[dict[str, Any]], lib.get_rows())
    row = next(r for r in rows if r["Product ID"] == "P-201")
    expected = SHEET1_ROWS[1]
    assert row["Description"] == expected["Description"]
    assert row["Price"] == expected["Price"]


@pytest.mark.parametrize("backend_name", EDITABLE_BACKENDS, ids=EDITABLE_BACKENDS)
def test_update_values_no_match_returns_zero_and_keeps_data_for_editable_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    rows_before = cast(list[dict[str, Any]], lib.get_rows())
    count = lib.update_values(
        search_criteria={"Product ID": "NOT-EXISTING"},
        values={"Location": "Should Not Be Applied"},
    )

    assert count == 0
    assert lib.get_rows() == rows_before


@pytest.mark.parametrize("backend_name", EDITABLE_BACKENDS, ids=EDITABLE_BACKENDS)
def test_update_values_partial_match_behaves_consistently_for_editable_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    count = lib.update_values(
        search_criteria={"Description": "Keyboard"},
        values={"Location": "Partially Matched"},
        partial_match=True,
    )

    assert count == 1
    rows = cast(list[dict[str, Any]], lib.get_rows())
    row = next(r for r in rows if r["Product ID"] == "P-201")
    assert row["Location"] == "Partially Matched"


@pytest.mark.parametrize("backend_name", EDITABLE_BACKENDS, ids=EDITABLE_BACKENDS)
def test_update_values_with_first_only_updates_only_one_of_multiple_matches(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    lib.update_values(search_criteria={"Product ID": "P-201"}, values={"Location": "Shared"})
    lib.update_values(search_criteria={"Product ID": "P-202"}, values={"Location": "Shared"})

    count = lib.update_values(
        search_criteria={"Location": "Shared"},
        values={"Description": "First Only Updated"},
        first_only=True,
    )

    assert count == 1
    rows = cast(list[dict[str, Any]], lib.get_rows())
    shared_rows = [r for r in rows if r["Location"] == "Shared"]
    assert len(shared_rows) == 2
    assert sum(1 for r in shared_rows if r["Description"] == "First Only Updated") == 1


@pytest.mark.parametrize("backend_name", EDITABLE_BACKENDS, ids=EDITABLE_BACKENDS)
def test_update_values_accepts_string_search_criteria_for_editable_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    count = lib.update_values(
        search_criteria="Product ID=P-203",
        values={"Location": "Archived"},
    )

    assert count == 1
    rows = cast(list[dict[str, Any]], lib.get_rows())
    row = next(r for r in rows if r["Product ID"] == "P-203")
    assert row["Location"] == "Archived"


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_update_values_header_row_out_of_range_raises_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)

    with pytest.raises(HeadersNotDeterminedException):
        lib.update_values(
            search_criteria={"Product ID": "P-200"},
            values={"Location": "Never Applied"},
            header_row=9999,
        )


@pytest.mark.parametrize("backend_name", STREAMING_BACKENDS, ids=STREAMING_BACKENDS)
def test_update_values_raises_in_read_only_mode_for_streaming_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)

    with pytest.raises(NullComponentException):
        lib.update_values(
            search_criteria={"Product ID": "P-200"},
            values={"Location": "Never Applied"},
        )


def test_update_values_raises_when_no_workbook_is_loaded(lib: RFExcelLibrary) -> None:
    with pytest.raises(WorkbookNotOpenException):
        lib.update_values(
            search_criteria={"Product ID": "P-200"},
            values={"Location": "Never Applied"},
        )
