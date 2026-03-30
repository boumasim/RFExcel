from pathlib import Path
from typing import Any, cast

import pytest

from rfexcel.exception.library_exceptions import (
    FileDoesNotExistException, HeadersNotDeterminedException,
    StreamingViolationException, WorkbookNotOpenException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import (BACKEND_NAMES, CSV_EDIT, CSV_HEADERS, CSV_ROWS, CSV_STREAM, EDITABLE_FORMAT_LIST,
                                  XLS_EDIT, XLS_HEADERS, XLS_ON_DEMAND,
                                  XLS_ROWS, XLSX_EDIT, XLSX_HEADERS, XLSX_ROWS,
                                  XLSX_STREAM, open_backend)

STREAMING_VIOLATION_BACKENDS = [XLSX_STREAM, CSV_STREAM]

EXPECTED_ROWS_BY_BACKEND: dict[str, list[dict[str, Any]]] = {
    XLSX_EDIT: XLSX_ROWS,
    XLSX_STREAM: XLSX_ROWS,
    CSV_EDIT: CSV_ROWS,
    CSV_STREAM: CSV_ROWS,
    XLS_EDIT: XLS_ROWS,
    XLS_ON_DEMAND: XLS_ROWS,
}

EXPECTED_HEADERS_BY_BACKEND: dict[str, list[str]] = {
    XLSX_EDIT: XLSX_HEADERS,
    XLSX_STREAM: XLSX_HEADERS,
    CSV_EDIT: CSV_HEADERS,
    CSV_STREAM: CSV_HEADERS,
    XLS_EDIT: XLS_HEADERS,
    XLS_ON_DEMAND: XLS_HEADERS,
}

EXACT_SEARCH_BY_BACKEND: dict[str, tuple[dict[str, Any], dict[str, Any]]] = {
    XLSX_EDIT: ({"Product ID": "P-202"}, XLSX_ROWS[2]),
    XLSX_STREAM: ({"Product ID": "P-202"}, XLSX_ROWS[2]),
    CSV_EDIT: ({"Product ID": "P-202"}, CSV_ROWS[2]),
    CSV_STREAM: ({"Product ID": "P-202"}, CSV_ROWS[2]),
    XLS_EDIT: ({"First Name": "Dulce"}, XLS_ROWS[0]),
    XLS_ON_DEMAND: ({"First Name": "Dulce"}, XLS_ROWS[0]),
}

PARTIAL_SEARCH_BY_BACKEND: dict[str, tuple[dict[str, Any], int]] = {
    XLSX_EDIT: ({"Description": "Keyboard"}, 1),
    XLSX_STREAM: ({"Description": "Keyboard"}, 1),
    CSV_EDIT: ({"Description": "Keyboard"}, 1),
    CSV_STREAM: ({"Description": "Keyboard"}, 1),
    XLS_EDIT: ({"Country": "United"}, 6),
    XLS_ON_DEMAND: ({"Country": "United"}, 6),
}

NO_MATCH_SEARCH_BY_BACKEND: dict[str, dict[str, Any]] = {
    XLSX_EDIT: {"Product ID": "NOPE"},
    XLSX_STREAM: {"Product ID": "NOPE"},
    CSV_EDIT: {"Product ID": "NOPE"},
    CSV_STREAM: {"Product ID": "NOPE"},
    XLS_EDIT: {"First Name": "NOPE"},
    XLS_ON_DEMAND: {"First Name": "NOPE"},
}


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_get_rows_returns_expected_data_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    assert lib.get_rows() == EXPECTED_ROWS_BY_BACKEND[backend_name]


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_get_rows_count_is_correct_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    assert len(lib.get_rows()) == len(EXPECTED_ROWS_BY_BACKEND[backend_name])


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_all_rows_have_expected_header_keys_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    expected_headers = EXPECTED_HEADERS_BY_BACKEND[backend_name]
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
    criteria, expected_value = EXACT_SEARCH_BY_BACKEND[backend_name]
    rows = lib.get_rows(search_criteria=criteria)
    assert len(rows) == 1
    assert rows[0] == expected_value


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_partial_match_search_behaves_consistently_across_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    criteria, expected_count = PARTIAL_SEARCH_BY_BACKEND[backend_name]
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
    assert result == EXPECTED_ROWS_BY_BACKEND[backend_name][0]


@pytest.mark.parametrize("backend_name", BACKEND_NAMES, ids=BACKEND_NAMES)
def test_one_row_no_match_returns_empty_dict_for_all_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)
    result = lib.get_rows(search_criteria=NO_MATCH_SEARCH_BY_BACKEND[backend_name], one_row=True)
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
