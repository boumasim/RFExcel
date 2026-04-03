from typing import Any

import pytest

from rfexcel.exception.library_exceptions import (NotMatchingColumns,
                                                  SheetDoesNotExistException,
                                                  StreamingViolationException,
                                                  WorkbookNotOpenException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import XLSX2_FILE
from tests.pyth.test_data import (BACKEND_NAMES, BACKENDS, EDITABLE_BACKENDS,
                                  STREAMING_BACKENDS, XLS_ON_DEMAND, XLSX_EDIT,
                                  XLSX_STREAM, open_backend)

SHEET_BACKENDS = [XLSX_EDIT, XLSX_STREAM, XLS_ON_DEMAND, XLSX_EDIT]
XLSX2_ONE_DIFF_SHEET = "OneDiff"
XLSX2_MORE_DIFFS_SHEET = "MoreDiffs"
XLSX2_DIFFS_AND_ROW_NUMBER_DIFF_SHEET = "DiffsAndRowsNumberDiff"
XLSX2_ROWS_NUMBER_DIFF_SHEET = "RowsNumberDiff"

BACKENDS_VS_XLSX2_ONEDIFF_DIFFS: list[dict[str, Any]] = [
    {
        "source_row_index": 5,
        "target_row_index": 5,
        "differences": {
            "Product ID": {"source": "P-203", "target": "P-205"},
            "Price": {"source": 5.99, "target": 6},
        },
    },
]

BACKENDS_VS_XLSX2_MORE_DIFFS: list[dict[str, Any]] = [
    {
        "source_row_index": 2,
        "target_row_index": 2,
        "differences": {
            "Product ID": {"source": "P-200", "target": "P-300"},
        },
    },
    {
        "source_row_index": 3,
        "target_row_index": 3,
        "differences": {
            "Product ID": {"source": "P-201", "target": "P-301"},
        },
    },
    {
        "source_row_index": 4,
        "target_row_index": 4,
        "differences": {
            "Product ID": {"source": "P-202", "target": "P-302"},
        },
    },
    {
        "source_row_index": 5,
        "target_row_index": 5,
        "differences": {
            "Product ID": {"source": "P-203", "target": "P-303"},
            "Description": {"source": "USB Cable", "target": "Cable"}
        },
    },
]

@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_compare_data_to_xlsx2_matches_expected_structure_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    assert lib.compare_data_to(XLSX2_FILE) == BACKENDS_VS_XLSX2_ONEDIFF_DIFFS


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_compare_data_to_xlsx2_one_diff_sheet_matches_expected_structure_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    assert lib.compare_data_to(XLSX2_FILE, target_sheet=XLSX2_ONE_DIFF_SHEET) == BACKENDS_VS_XLSX2_ONEDIFF_DIFFS


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_compare_data_to_xlsx2_more_diffs_sheet_matches_expected_structure_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    assert lib.compare_data_to(XLSX2_FILE, target_sheet=XLSX2_MORE_DIFFS_SHEET) == BACKENDS_VS_XLSX2_MORE_DIFFS


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_compare_data_to_xlsx2_with_header_subset_reports_only_requested_columns_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    result = lib.compare_data_to(XLSX2_FILE, headers=["Product ID"])
    assert len(result) == 1
    assert "Product ID" in result[0]["differences"]
    assert "Price" not in result[0]["differences"]


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_compare_data_to_xlsx2_with_identical_header_filter_returns_empty_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    assert lib.compare_data_to(XLSX2_FILE, headers=["Description"]) == []


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_compare_data_to_default_target_sheet_matches_explicit_first_sheet_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    default_result = lib.compare_data_to(XLSX2_FILE)

    path, read_only = BACKENDS[backend_name]
    lib.load_workbook(path, read_only=read_only)
    explicit_result = lib.compare_data_to(XLSX2_FILE, target_sheet=XLSX2_ONE_DIFF_SHEET)

    assert default_result == explicit_result


@pytest.mark.parametrize("backend_name", SHEET_BACKENDS, ids=SHEET_BACKENDS)
def test_compare_data_to_target_sheet_with_source_on_sheet_two_raises_for_xlsx_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    lib.switch_sheet("Sheet2")
    with pytest.raises(NotMatchingColumns):
        lib.compare_data_to(XLSX2_FILE, target_sheet=XLSX2_ONE_DIFF_SHEET)


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_compare_data_to_xlsx2_more_diffs_sheet_returns_four_differences_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    result = lib.compare_data_to(XLSX2_FILE, target_sheet=XLSX2_MORE_DIFFS_SHEET)
    assert len(result) == 4


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_compare_data_to_xlsx2_diffs_and_row_number_diff_sheet_returns_empty_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    assert lib.compare_data_to(XLSX2_FILE, target_sheet=XLSX2_DIFFS_AND_ROW_NUMBER_DIFF_SHEET) == BACKENDS_VS_XLSX2_MORE_DIFFS


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_compare_data_to_xlsx2_rows_number_diff_sheet_returns_empty_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    assert lib.compare_data_to(XLSX2_FILE, target_sheet=XLSX2_ROWS_NUMBER_DIFF_SHEET) == []


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_compare_data_to_itself_returns_no_differences_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    assert lib.compare_data_to(BACKENDS[backend_name][0]) == []


def test_compare_data_to_raises_when_no_workbook_is_loaded(lib: RFExcelLibrary) -> None:
    with pytest.raises(WorkbookNotOpenException):
        lib.compare_data_to(XLSX2_FILE)


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_compare_data_to_raises_for_unknown_header_filter_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    with pytest.raises(NotMatchingColumns):
        lib.compare_data_to(XLSX2_FILE, headers=["Nonexistent"])


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_compare_data_to_same_path_returns_no_differences_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    source_path, _ = BACKENDS[backend_name]
    assert lib.compare_data_to(source_path) == []


@pytest.mark.parametrize("backend_name", EDITABLE_BACKENDS, ids=EDITABLE_BACKENDS)
def test_compare_data_to_same_path_keeps_workbook_usable_for_editable_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    source_path, _ = BACKENDS[backend_name]
    lib.compare_data_to(source_path)

    rows = lib.get_rows()
    assert len(rows) > 0


@pytest.mark.parametrize("backend_name", STREAMING_BACKENDS, ids=STREAMING_BACKENDS)
def test_compare_data_to_same_path_followed_by_get_rows_raises_for_streaming_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    source_path, _ = BACKENDS[backend_name]
    lib.compare_data_to(source_path)

    with pytest.raises(StreamingViolationException):
        lib.get_rows()

@pytest.mark.parametrize("backend_name", SHEET_BACKENDS, ids=SHEET_BACKENDS)
def test_compare_data_to_same_workbook_different_sheet_raises_sheet_not_found(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    source_path, _ = BACKENDS[backend_name]
    with pytest.raises(SheetDoesNotExistException):
        lib.compare_data_to(source_path, target_sheet="Nope")


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_compare_data_to_fail_on_diff_true_with_same_file_does_not_raise_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    source_path, _ = BACKENDS[backend_name]
    result = lib.compare_data_to(source_path, fail_on_diff=True)
    assert result == []


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_compare_data_to_fail_on_diff_true_raises_assertion_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    with pytest.raises(AssertionError):
        lib.compare_data_to(XLSX2_FILE, fail_on_diff=True)


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
@pytest.mark.parametrize(
    "target_sheet",
    [XLSX2_DIFFS_AND_ROW_NUMBER_DIFF_SHEET, XLSX2_ROWS_NUMBER_DIFF_SHEET],
    ids=["diffs_and_row_number_diff", "rows_number_diff"],
)
def test_compare_data_to_row_count_mismatch_sheets_raise_assertion_with_fail_on_diff_true_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    target_sheet: str,
) -> None:
    open_backend(lib, backend_name)
    with pytest.raises(AssertionError):
        lib.compare_data_to(XLSX2_FILE, target_sheet=target_sheet, fail_on_diff=True)


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_compare_data_to_fail_on_diff_respects_headers_filter_and_raises_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    with pytest.raises(AssertionError):
        lib.compare_data_to(XLSX2_FILE, headers=["Product ID"], fail_on_diff=True)
