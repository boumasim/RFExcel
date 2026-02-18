from itertools import zip_longest
from typing import Any, List, Optional, override

from openpyxl import Workbook
from openpyxl.chartsheet import Chartsheet
from openpyxl.worksheet.worksheet import Worksheet

from rfexcel.exception.library_exceptions import (LibraryException,
                                                  StreamingViolationException)
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.model.raw_data.xlsx_raw_row_data import XlsxRawRowData
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
    def fetch_row(self, row_index: int) -> Row:
        if not self._active_sheet:
            raise LibraryException("No active worksheet")
        
        target_excel_row = row_index
        
        if target_excel_row > self._active_sheet.max_row:
            raise StopIteration()
        
        row_values = next(
            self._active_sheet.iter_rows(min_row=target_excel_row, max_row=target_excel_row, values_only=True)
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
    def __init__(self, wb: Workbook):
        self._wb = wb
        self._active_sheet = self._wb.active
        self._row_generator = self._active_sheet.iter_rows(values_only=True) if self._active_sheet else iter([])
        self._last_read_row_index = 0

    @property
    @override
    def get_active_sheet(self) -> Worksheet | Chartsheet | None:
        return self._active_sheet
    
    @property
    @override
    def last_read_row_index(self) -> int:
        return self._last_read_row_index
    
    def fetch_row(self, row_index: int) -> IRawRowData:
        row_data = next(self._row_generator)
        self._last_read_row_index = self._last_read_row_index + 1
        return XlsxRawRowData(row_data,True)

    @override
    def close(self):
        self._wb.close()