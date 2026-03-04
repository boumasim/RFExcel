from collections.abc import Iterator
from pathlib import Path
from typing import Any, override

from openpyxl import Workbook
from openpyxl.chartsheet import Chartsheet
from openpyxl.worksheet.worksheet import Worksheet

from rfexcel.exception.library_exceptions import (LibraryException,
                                                  NotSupportedInReadOnlyMode)
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.model.raw_data.xlsx_raw_row_data import XlsxRawRowData

from .i_resource import IResource


class XlsxEditResource(IResource):
    def __init__(self, wb: Workbook, path: Path):
        super().__init__(path)
        self._wb: Workbook = wb
        self._active_sheet: Worksheet | None = wb.worksheets[0] if wb.worksheets else None

    @property
    @override
    def get_active_sheet(self) -> Worksheet | Chartsheet | None:
        return self._active_sheet

    @property
    @override
    def last_read_row_index(self) -> int:
        return -1

    @override
    def fetch_row(self, row_index: int, **kwargs: Any) -> IRawRowData:
        if not self._active_sheet:
            raise LibraryException("No active worksheet")
        if row_index > self._active_sheet.max_row:
            raise StopIteration()
        data_only: bool = kwargs.get('data_only', True)  # type: ignore[assignment]
        row_values = next(
            self._active_sheet.iter_rows(min_row=row_index, max_row=row_index, values_only=data_only)
        )
        return XlsxRawRowData(row_values, data_only)

    @override
    def get_sheet_names(self) -> list[str]:
        return list(self._wb.sheetnames)

    @override
    def switch_sheet(self, name: str) -> None:
        self._active_sheet = self._wb[name]

    @override
    def add_sheet(self, name: str) -> None:
        ws: Worksheet = self._wb.create_sheet(title=name)
        self._active_sheet = ws

    @override
    def delete_sheet(self, name: str) -> None:
        if name not in self._wb.sheetnames:
            raise LibraryException(f"Sheet '{name}' does not exist")
        del self._wb[name]
        self._active_sheet = self._wb.worksheets[0] if self._wb.worksheets else None

    @override
    def close(self):
        self._wb.close()


class XlsxStreamResource(IResource):
    def __init__(self, wb: Workbook, path: Path):
        super().__init__(path)
        self._wb = wb
        self._active_sheet = self._wb.worksheets[0] if self._wb.worksheets else None
        self._row_generator: Iterator[tuple[Any, ...]] | None = None  # lazily initialised on first fetch_row call
        self._last_read_row_index = 0

    @property
    @override
    def get_active_sheet(self) -> Worksheet | Chartsheet | None:
        return self._active_sheet
    
    @property
    @override
    def last_read_row_index(self) -> int:
        return self._last_read_row_index
    
    @override
    def fetch_row(self, row_index: int, **kwargs: Any) -> IRawRowData:
        if self._row_generator is None:
            data_only: bool = kwargs.get('data_only', True)  # type: ignore[assignment]
            self._row_generator = (
                self._active_sheet.iter_rows(values_only=data_only)
                if self._active_sheet
                else iter([])
            )
        row_data = next(self._row_generator)
        self._last_read_row_index += 1
        return XlsxRawRowData(row_data, kwargs.get('data_only', True))  # type: ignore[arg-type]

    @override
    def get_sheet_names(self) -> list[str]:
        return list(self._wb.sheetnames)

    @override
    def switch_sheet(self, name: str) -> None:
        self._active_sheet = self._wb[name]
        self._row_generator = None
        self._last_read_row_index = 0

    @override
    def add_sheet(self, name: str) -> None:
        raise NotSupportedInReadOnlyMode("Adding sheets is not supported in streaming mode")

    @override
    def delete_sheet(self, name: str) -> None:
        raise NotSupportedInReadOnlyMode("Deleting sheets is not supported in streaming mode")

    @override
    def close(self):
        self._wb.close()