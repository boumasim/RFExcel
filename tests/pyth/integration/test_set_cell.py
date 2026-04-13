from pathlib import Path

import pytest

from rfexcel.exception.library_exceptions import (
	InvalidCellNameException,
	NullComponentException,
	OperationNotSupportedForFormat,
)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import (
	CSV_EDIT,
	STREAMING_BACKENDS,
	XLS_EDIT,
	XLSX_EDIT,
	load_backend_copy,
	open_backend,
)

EDIT_BACKENDS = [XLSX_EDIT, XLS_EDIT]

@pytest.mark.parametrize("backend_name", EDIT_BACKENDS, ids=EDIT_BACKENDS)
def test_set_cell_writes_string_value_for_edit_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    lib.set_cell("A2", "NEW_VALUE")

    assert lib.get_cell("A2") == "NEW_VALUE"


@pytest.mark.parametrize("backend_name", EDIT_BACKENDS, ids=EDIT_BACKENDS)
def test_set_cell_writes_int_value_for_edit_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    lib.set_cell("C2", 999)

    assert lib.get_cell("C2") == 999


@pytest.mark.parametrize("backend_name", EDIT_BACKENDS, ids=EDIT_BACKENDS)
def test_set_cell_writes_float_value_for_edit_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    lib.set_cell("C2", 3.14)

    assert lib.get_cell("C2") == 3.14


@pytest.mark.parametrize("backend_name", EDIT_BACKENDS, ids=EDIT_BACKENDS)
def test_set_cell_writes_bool_value_for_edit_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    lib.set_cell("B2", True)

    assert lib.get_cell("B2") is True


@pytest.mark.parametrize("backend_name", EDIT_BACKENDS, ids=EDIT_BACKENDS)
def test_set_cell_overwrites_existing_value_for_edit_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    lib.set_cell("A2", "OVERWRITTEN")

    assert lib.get_cell("A2") == "OVERWRITTEN"


@pytest.mark.parametrize("backend_name", EDIT_BACKENDS, ids=EDIT_BACKENDS)
def test_set_cell_does_not_affect_other_cells_for_edit_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    lib.set_cell("A2", "CHANGED")

    assert lib.get_cell("B2") == "Wireless Mouse"


@pytest.mark.parametrize("backend_name", EDIT_BACKENDS, ids=EDIT_BACKENDS)
def test_set_cell_raises_for_invalid_cell_name_on_edit_backends(
    lib: RFExcelLibrary,
    backend_name: str,
    tmp_path: Path,
) -> None:
    load_backend_copy(lib, backend_name, tmp_path)

    with pytest.raises(InvalidCellNameException):
        lib.set_cell("not-a-cell", "value")


@pytest.mark.parametrize("backend_name", STREAMING_BACKENDS, ids=STREAMING_BACKENDS)
def test_set_cell_raises_for_streaming_backends(
    lib: RFExcelLibrary,
    backend_name: str,
) -> None:
    open_backend(lib, backend_name)

    with pytest.raises(NullComponentException):
        lib.set_cell("A2", "value")


def test_set_cell_raises_for_csv_edit(
    lib: RFExcelLibrary,
) -> None:
    open_backend(lib, CSV_EDIT)

    with pytest.raises(OperationNotSupportedForFormat):
        lib.set_cell("A2", "value")
