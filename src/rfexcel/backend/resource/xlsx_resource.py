from collections.abc import Iterator
from typing import Any, override

from openpyxl import Workbook
from openpyxl.chartsheet import Chartsheet
from openpyxl.worksheet.worksheet import Worksheet

from rfexcel.exception.library_exceptions import LibraryException
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.model.raw_data.xlsx_raw_row_data import XlsxRawRowData

from .i_resource import IResource


class XlsxEditResource(IResource):
    def __init__(self, wb: Workbook):
        self._wb: Workbook = wb
        self._active_sheet: Worksheet | None = wb.active

    @property
    @override
    def get_active_sheet(self) -> Worksheet | Chartsheet | None:
        return self._active_sheet

    @property
    @override
    def last_read_row_index(self) -> int:
        return -1

    @override
    def fetch_row(self, row_index: int, data_only: bool = True) -> IRawRowData:
        if not self._active_sheet:
            raise LibraryException("No active worksheet")
        if row_index > self._active_sheet.max_row:
            raise StopIteration()
        row_values = next(
            self._active_sheet.iter_rows(min_row=row_index, max_row=row_index, values_only=data_only)
        )
        return XlsxRawRowData(row_values, data_only)

    @override
    def close(self):
        self._wb.close()


class XlsxStreamResource(IResource):
    def __init__(self, wb: Workbook):
        self._wb = wb
        self._active_sheet = self._wb.active
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
    def fetch_row(self, row_index: int, data_only: bool = True) -> IRawRowData:
        if self._row_generator is None:
            self._row_generator = (
                self._active_sheet.iter_rows(values_only=data_only)
                if self._active_sheet
                else iter([])
            )
        row_data = next(self._row_generator)
        self._last_read_row_index += 1
        return XlsxRawRowData(row_data, data_only)

    @override
    def close(self):
        self._wb.close()