import pytest

from rfexcel.exception.library_exceptions import (
    InvalidCellNameException, OperationNotSupportedForFormat,
    StreamingViolationException)
from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.test_data import (CSV_EDIT, CSV_STREAM, XLS_EDIT,
                                  XLS_ON_DEMAND, XLSX_EDIT, XLSX_STREAM,
                                  open_backend)

SUPPORTED_BACKENDS = [XLSX_EDIT, XLSX_STREAM, XLS_EDIT, XLS_ON_DEMAND]
CSV_BACKENDS = [CSV_EDIT, CSV_STREAM]
STREAMING_BACKENDS = [XLSX_STREAM, XLS_ON_DEMAND]


@pytest.mark.parametrize("backend_name", SUPPORTED_BACKENDS, ids=SUPPORTED_BACKENDS)
def test_get_cell_returns_expected_value_for_supported_backends(
	lib: RFExcelLibrary,
	backend_name: str,
) -> None:
	open_backend(lib, backend_name)

	assert lib.get_cell("A2") == "P-200"
	if backend_name in [XLSX_STREAM, XLS_ON_DEMAND]:
		assert lib.get_cell("C3") == 89.99
	else:
		assert lib.get_cell("C2") == 25.5
	assert lib.get_cell("C4") == 150


@pytest.mark.parametrize("backend_name", SUPPORTED_BACKENDS, ids=SUPPORTED_BACKENDS)
def test_get_cell_returns_empty_string_for_blank_supported_cells(
	lib: RFExcelLibrary,
	backend_name: str,
) -> None:
	open_backend(lib, backend_name)

	assert lib.get_cell("Z999") == ""


@pytest.mark.parametrize("backend_name", CSV_BACKENDS, ids=CSV_BACKENDS)
def test_get_cell_raises_for_csv_backends(
	lib: RFExcelLibrary,
	backend_name: str,
) -> None:
	open_backend(lib, backend_name)

	with pytest.raises(OperationNotSupportedForFormat):
		lib.get_cell("A2")


@pytest.mark.parametrize("backend_name", SUPPORTED_BACKENDS, ids=SUPPORTED_BACKENDS)
def test_get_cell_raises_for_invalid_coordinate_on_supported_backends(
	lib: RFExcelLibrary,
	backend_name: str,
) -> None:
	open_backend(lib, backend_name)

	with pytest.raises(InvalidCellNameException):
		lib.get_cell("not-a-cell")

@pytest.mark.parametrize("backend_name", STREAMING_BACKENDS, ids=STREAMING_BACKENDS)
def test_get_cell_raises_when_revisiting_same_row_in_xlsx_stream_mode(
	lib: RFExcelLibrary,
	backend_name: str,
) -> None:
	open_backend(lib, backend_name)
	assert lib.get_cell("A5") == "P-203"

	with pytest.raises(StreamingViolationException):
		lib.get_cell("B5")