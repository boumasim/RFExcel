from pathlib import Path
from typing import Any, cast

import pytest

from rfexcel.exception.library_exceptions import (
	FileDoesNotExistException,
	HeadersNotDeterminedException,
	StreamingViolationException,
	WorkbookNotOpenException,
)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import (
	BACKEND_NAMES,
	CSV_STREAM,
	EDITABLE_FORMAT_LIST,
	SHEET1_HEADERS,
	SHEET1_ROWS,
	XLSX_EDIT,
	XLSX_STREAM,
	open_backend,
)

STREAMING_VIOLATION_BACKENDS = [XLSX_STREAM, CSV_STREAM]

EXACT_SEARCH: tuple[dict[str, Any], dict[str, Any]] = ({"Product ID": "P-202"}, SHEET1_ROWS[2])

PARTIAL_SEARCH: tuple[dict[str, Any], int] = ({"Description": "Keyboard"}, 1)

NO_MATCH_SEARCH: dict[str, Any] = {"Product ID": "NOPE"}

@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_get_rows_returns_expected_data_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    assert lib.get_rows() == SHEET1_ROWS


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_get_rows_count_is_correct_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    assert len(lib.get_rows()) == len(SHEET1_ROWS)


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_all_rows_have_expected_header_keys_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    expected_headers = SHEET1_HEADERS
    for row in lib.get_rows():
        typed_row = cast(dict[str, Any], row)
        assert list(typed_row.keys()) == expected_headers


@pytest.mark.parametrize(
    "backend_name",
    BACKEND_NAMES,
    ids=BACKEND_NAMES,
)
def test_header_row_out_of_range_raises_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    with pytest.raises(HeadersNotDeterminedException):
        lib.get_rows(header_row=9999)


@pytest.mark.parametrize("backend_name", STREAMING_VIOLATION_BACKENDS, ids=STREAMING_VIOLATION_BACKENDS)
def test_streaming_backends_raise_when_get_rows_called_twice(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    lib.get_rows()
    with pytest.raises(StreamingViolationException):
        lib.get_rows()


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_exact_search_criteria_returns_single_expected_match(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    criteria, expected_value = EXACT_SEARCH
    rows = lib.get_rows(search_criteria=criteria)
    assert len(rows) == 1
    assert rows[0] == expected_value


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_partial_match_search_behaves_consistently_across_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    criteria, expected_count = PARTIAL_SEARCH
    rows = lib.get_rows(search_criteria=criteria, partial_match=True)
    assert len(rows) == expected_count


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_one_row_true_returns_first_row_dict_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    result = lib.get_rows(one_row=True)
    assert isinstance(result, dict)
    assert result == SHEET1_ROWS[0]


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_one_row_no_match_returns_empty_dict_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    result = lib.get_rows(search_criteria=NO_MATCH_SEARCH, one_row=True)
    assert result == {}

@pytest.mark.parametrize("format_name", EDITABLE_FORMAT_LIST, ids=EDITABLE_FORMAT_LIST)
def test_get_rows_on_empty_file_raises(
    lib: RFExcelLibrary,
    format_name: str,
    tmp_path: Path
) -> None:
    lib.create_workbook(str(tmp_path / f"empty.{format_name}"))
    with pytest.raises(HeadersNotDeterminedException):
        lib.get_rows()


def test_get_rows_raises_when_no_workbook_is_loaded(lib: RFExcelLibrary) -> None:
    with pytest.raises(WorkbookNotOpenException):
        lib.get_rows()


def test_get_rows_raises_after_close(lib: RFExcelLibrary) -> None:
    open_backend(lib, XLSX_EDIT)
    lib.close()
    with pytest.raises(WorkbookNotOpenException):
        lib.get_rows()


def test_load_nonexistent_file_raises(lib: RFExcelLibrary) -> None:
    with pytest.raises(FileDoesNotExistException):
        lib.load_workbook("/nonexistent/path/missing.xlsx")
