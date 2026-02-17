from itertools import zip_longest
from typing import Any, List, Optional, override

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from rfexcel.exception.library_exceptions import (LibraryException,
                                                  RowIndexOutOfBoundsException,
                                                  StreamingViolationException)
from rfexcel.utlis.types import Row

from .i_resource import IResource


class XlsxEditResource(IResource):
    def __init__(self, wb: Workbook, header_row: int = 1):
        self._wb: Workbook = wb
        self._header_row = header_row
        self._active_sheet: Worksheet | None = wb.active
        
        self._headers = self._load_headers()
    
    def _load_headers(self) -> list[str]:
        if not self._active_sheet:
            return []
        
        header_generator = self._active_sheet.iter_rows(
            min_row=self._header_row,
            max_row=self._header_row,
            values_only=True
        )
        
        header_data = next(header_generator, None)
        if header_data:
            return [str(cell) if cell is not None else "" for cell in header_data]
        return []

    @property
    @override
    def header_row(self) -> int:
        return self._header_row

    @override
    def get_row(self, row_index: int) -> Row:
        if not self._active_sheet:
            raise LibraryException("No active worksheet")
        
        target_excel_row = row_index
        
        if target_excel_row > self._active_sheet.max_row:
            raise StopIteration()
        
        row_values = next(
            self._active_sheet.iter_rows(min_row=target_excel_row, max_row=target_excel_row, values_only=True),
            None
        )
        
        if row_values is None:
            raise StopIteration()

        return {
            header: str(val) if val is not None else ""
            for header, val in zip_longest(self._headers, row_values, fillvalue=None)
            if header is not None
        }

    @override
    def close(self):
        self._wb.close()


class XlsxStreamResource(IResource):
    def __init__(self, wb: Workbook, header_row: int = 1):
        self._wb = wb
        self._header_row = header_row
        self._active_sheet = wb.active
        
        self._row_generator = self._active_sheet.iter_rows(values_only=True) if self._active_sheet else iter([])
        self._headers = self._load_headers()
        self._last_read_data_index = self.header_row
    
    def _load_headers(self) -> list[str]:
        if not self._active_sheet:
            return []
        
        for i in range(self._header_row):
            try:
                row_data = next(self._row_generator)
                if i == self._header_row - 1:
                    return [str(cell) if cell is not None else "" for cell in row_data]
            except StopIteration:
                break
        return []

    @property
    @override
    def header_row(self) -> int:
        return self._header_row

    @override
    def get_row(self, row_index: int) -> Row:
        if row_index <= self._last_read_data_index:
            raise StreamingViolationException(row_index, self._last_read_data_index)

        row_data = None
        
        while self._last_read_data_index < row_index:
            try:
                row_data = next(self._row_generator)
                self._last_read_data_index += 1
            except StopIteration:
                raise StopIteration()
        
        if row_data is None:
            raise StopIteration()

        return dict(zip_longest(self._headers, (str(v) if v is not None else "" for v in row_data), fillvalue=""))

    @override
    def close(self):
        self._wb.close()