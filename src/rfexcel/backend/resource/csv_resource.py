import csv
from pathlib import Path
from typing import Any, override

from robot.api import logger

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.exception.library_exceptions import OperationNotSupportedForFormat
from rfexcel.model.raw_data.csv_raw_row_data import CsvRawRowData
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.rfexcel_constants import (BASE_DIALECT, BASE_ENCODING,
                                       CSV_NOT_SUPPORTED_MSG)


class CsvEditResource(IResource):
    def __init__(self, path: Path, dialect: str = BASE_DIALECT, encoding: str = BASE_ENCODING, **kwargs: Any):
        super().__init__(path)
        self._encoding = encoding
        self._dialect = dialect
        self._edited = False

        with open(path, mode='r', newline='', encoding=encoding) as f:
            self._all_rows: list[list[str]] = list(csv.reader(f, dialect=dialect, **kwargs))

    @property
    @override
    def get_active_sheet(self) -> None:
        return None

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
    def close(self):
        if self._edited:
            with open(self._path, mode='w', newline='', encoding=self._encoding) as f:
                writer = csv.writer(f, dialect=self._dialect)
                writer.writerows(self._all_rows)
            logger.info(f"CSV file '{self._path.name}' was updated.")


class CsvStreamResource(IResource):
    def __init__(self, path: Path, dialect: str = BASE_DIALECT, encoding: str = BASE_ENCODING, **kwargs: Any):
        super().__init__(path)
        self._handle = open(path, mode='r', newline='', encoding=encoding)
        self._reader = csv.reader(self._handle, dialect=dialect, **kwargs)
        self._last_read_row_index: int = 0

    @property
    @override
    def get_active_sheet(self) -> None:
        return None

    @property
    @override
    def last_read_row_index(self) -> int:
        return self._last_read_row_index

    @override
    def fetch_row(self, row_index: int, **kwargs: Any) -> IRawRowData:
        try:
            raw_row = next(self._reader)
        except StopIteration:
            raise StopIteration()
        self._last_read_row_index += 1
        return CsvRawRowData(raw_row)

    @override
    def get_sheet_names(self) -> list[str]:
        raise OperationNotSupportedForFormat()

    @override
    def switch_sheet(self, name: str) -> None:
        raise OperationNotSupportedForFormat()

    @override
    def add_sheet(self, name: str) -> None:
        raise OperationNotSupportedForFormat("CSV files do not support multiple sheets")

    @override
    def delete_sheet(self, name: str) -> None:
        raise OperationNotSupportedForFormat("CSV files do not support multiple sheets")

    @override
    def close(self):
        if self._handle and not self._handle.closed:
            self._handle.close()