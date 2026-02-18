from itertools import zip_longest
from typing import override

import xlrd.sheet
from xlrd import Book

from rfexcel.exception.library_exceptions import (LibraryException,
                                                  StreamingViolationException)
from rfexcel.utlis.types import Row

from .i_resource import IResource


class XlsEditResource(IResource):
    def __init__(self, wb: Book, header_row: int = 1):
        self._wb: Book = wb
        self._header_row = header_row
        self._active_sheet: xlrd.sheet.Sheet | None = wb.sheet_by_index(0) if wb.nsheets > 0 else None
        
        self._headers = self._load_headers()
    
    def _load_headers(self) -> list[str]:
        if not self._active_sheet:
            return []
        
        header_idx = self._header_row - 1
        if 0 <= header_idx < self._active_sheet.nrows:
            return [str(v) for v in self._active_sheet.row_values(header_idx)]
        return []

    @property
    @override
    def header_row(self) -> int:
        return self._header_row

    @override
    def fetch_row(self, row_index: int) -> Row:
        if not self._active_sheet:
            raise LibraryException("No active worksheet")
        
        target_xlrd_index = row_index - 1
        
        if target_xlrd_index >= self._active_sheet.nrows:
            raise StopIteration()
        
        row_data = self._active_sheet.row_values(target_xlrd_index)
        
        return {
            header: str(val) if val is not None else ""
            for header, val in zip_longest(self._headers, row_data, fillvalue="")
            if header is not None
        }

    @override
    def close(self):
        self._wb.release_resources()


class XlsStreamResource(IResource):
    def __init__(self, wb: Book, header_row: int = 1):
        self._wb: Book = wb
        self._header_row = header_row
        self._active_sheet: xlrd.sheet.Sheet | None = wb.sheet_by_index(0) if wb.nsheets > 0 else None
        
        self._headers = self._load_headers()
        self._last_read_data_index: int = 0
    
    def _load_headers(self) -> list[str]:
        if not self._active_sheet:
            return []
        
        header_idx = self._header_row - 1
        if 0 <= header_idx < self._active_sheet.nrows:
            return [str(v) for v in self._active_sheet.row_values(header_idx)]
        return []

    @property
    @override
    def header_row(self) -> int:
        return self._header_row

    @override
    def fetch_row(self, row_index: int) -> Row:
        if not self._active_sheet:
            raise LibraryException("No active worksheet")
        
        if row_index <= self._last_read_data_index:
            raise StreamingViolationException(row_index, self._last_read_data_index)
        
        self._last_read_data_index = row_index
        
        target_xlrd_index = row_index - 1
        
        if target_xlrd_index >= self._active_sheet.nrows:
            raise StopIteration()
        
        row_data = self._active_sheet.row_values(target_xlrd_index)
        
        return {
            header: str(val) if val is not None else ""
            for header, val in zip_longest(self._headers, row_data, fillvalue="")
            if header is not None
        }

    @override
    def close(self):
        self._wb.release_resources()