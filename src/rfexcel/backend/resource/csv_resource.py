import csv
from pathlib import Path
from typing import Any, override

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.exception.library_exceptions import FileSaveException, NotSupportedInReadOnlyMode, OperationNotSupportedForFormat
from rfexcel.model.cell_data.i_raw_cell_data import IRawCellData
from rfexcel.model.raw_data.csv_raw_row_data import CsvRawRowData
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.rfexcel_constants import BASE_DIALECT, BASE_ENCODING, CSV_NOT_SUPPORTED_MSG
from rfexcel.utils.library_logger import logger
from rfexcel.utils.types import ColumnValues, InsertNativeType


class CsvEditResource(IResource):
    def __init__(self, path: Path, dialect: str = BASE_DIALECT, encoding: str = BASE_ENCODING, **kwargs: Any):
        super().__init__(path)
        self._encoding = encoding
        self._dialect = dialect

        with open(path, mode='r', newline='', encoding=encoding) as f:
            self._all_rows: list[list[str]] = list(csv.reader(f, dialect=dialect, **kwargs))

    @property
    @override
    def active_sheets(self) -> None:
        return None
    
    @property
    @override
    def current_sheet(self) -> str:
        raise OperationNotSupportedForFormat("CSV files do not have sheets; current_sheet is not applicable")

    @property
    @override
    def last_read_row_index(self) -> int:
        return -1

    @override
    def fetch_row(self, row_index: int, **kwargs: Any) -> IRawRowData:
        list_index = row_index - 1

        if list_index < 0 or list_index >= len(self._all_rows):
            raise StopIteration()

        return CsvRawRowData(self._all_rows[list_index])

    @override
    def fetch_cell(self, cell_name: str, **kwargs: Any) -> IRawCellData:
        raise OperationNotSupportedForFormat("Get Cell is supported only for .xlsx and .xls files")

    @override
    def get_sheet_names(self) -> list[str]:
        raise OperationNotSupportedForFormat(CSV_NOT_SUPPORTED_MSG)

    @override
    def switch_sheet(self, name: str) -> None:
        raise OperationNotSupportedForFormat("This operation is not supported for CSV files")

    @override
    def add_sheet(self, name: str) -> None:
        raise OperationNotSupportedForFormat("CSV files do not support multiple sheets")

    @override
    def delete_sheet(self, name: str) -> None:
        raise OperationNotSupportedForFormat("CSV files do not support multiple sheets")

    @override
    def save(self, path: Path | None = None) -> None:
        target = path or self._path
        try:
            with open(target, mode='w', newline='', encoding=self._encoding) as f:
                writer = csv.writer(f, dialect=self._dialect)
                writer.writerows(self._all_rows)
        except Exception as e:
            raise FileSaveException(str(target), str(e)) from e
        self._path = target
        logger.info(f"CSV file saved to '{target.name}'.")

    @override
    def append_row(self, cell_data: ColumnValues) -> None:
        if not cell_data:
            self._all_rows.append([])
            return
        max_col = max(cell_data.keys())
        row = [cell_data.get(i, "") for i in range(1, max_col + 1)]
        self._all_rows.append([str(cell) for cell in row])

    @override
    def update_row(self, row_index: int, cell_data: ColumnValues) -> None:
        list_index = row_index - 1
        if list_index < 0 or list_index >= len(self._all_rows):
            return
        row = self._all_rows[list_index]
        for col, value in cell_data.items():
            col_index = col - 1
            while len(row) <= col_index:
                row.append("")
            row[col_index] = str(value)

    @override
    def delete_row(self, row_index: int) -> None:
        list_index = row_index - 1
        if 0 <= list_index < len(self._all_rows):
            self._all_rows.pop(list_index)

    @override
    def insert_row(self, row_index: int, cell_data: ColumnValues) -> None:
        list_index = row_index - 1
        if not cell_data:
            row: list[str] = []
        else:
            max_col = max(cell_data.keys())
            row = [str(cell_data.get(i, "")) for i in range(1, max_col + 1)]
        self._all_rows.insert(list_index, row)

    @override
    def set_cell(self, cell_name: str, value: InsertNativeType) -> None:
        raise OperationNotSupportedForFormat("Set Cell is supported only for .xlsx and .xls files")

    @override
    def close(self):
        pass


class CsvStreamResource(IResource):
    def __init__(self, path: Path, dialect: str = BASE_DIALECT, encoding: str = BASE_ENCODING, **kwargs: Any):
        super().__init__(path)
        self._handle = open(path, mode='r', newline='', encoding=encoding)
        self._reader = csv.reader(self._handle, dialect=dialect, **kwargs)
        self._last_read_row_index: int = 0

    @property
    @override
    def active_sheets(self) -> None:
        return None

    @property
    @override
    def current_sheet(self) -> str:
        raise OperationNotSupportedForFormat()

    @property
    @override
    def last_read_row_index(self) -> int:
        return self._last_read_row_index

    @override
    def fetch_row(self, row_index: int, **kwargs: Any) -> IRawRowData:
        while self._last_read_row_index < row_index - 1:
            next(self._reader)
            self._last_read_row_index += 1
        raw_row = next(self._reader)
        self._last_read_row_index += 1
        return CsvRawRowData(raw_row)

    @override
    def fetch_cell(self, cell_name: str, **kwargs: Any) -> IRawCellData:
        raise OperationNotSupportedForFormat("Get Cell is supported only for .xlsx and .xls files")

    @override
    def get_sheet_names(self) -> list[str]:
        raise OperationNotSupportedForFormat()

    @override
    def switch_sheet(self, name: str) -> None:
        raise OperationNotSupportedForFormat()

    @override
    def add_sheet(self, name: str) -> None:
        raise OperationNotSupportedForFormat()

    @override
    def delete_sheet(self, name: str) -> None:
        raise OperationNotSupportedForFormat()

    @override
    def save(self, path: Path | None = None) -> None:
        raise NotSupportedInReadOnlyMode()

    @override
    def append_row(self, cell_data: ColumnValues) -> None:
        raise NotSupportedInReadOnlyMode()
    @override
    def update_row(self, row_index: int, cell_data: ColumnValues) -> None:
        raise NotSupportedInReadOnlyMode()
    @override
    def delete_row(self, row_index: int) -> None:
        raise NotSupportedInReadOnlyMode()

    @override
    def insert_row(self, row_index: int, cell_data: ColumnValues) -> None:
        raise NotSupportedInReadOnlyMode()

    @override
    def set_cell(self, cell_name: str, value: InsertNativeType) -> None:
        raise OperationNotSupportedForFormat("Set Cell is supported only for .xlsx and .xls files")

    @override
    def close(self):
        if self._handle and not self._handle.closed:
            self._handle.close()