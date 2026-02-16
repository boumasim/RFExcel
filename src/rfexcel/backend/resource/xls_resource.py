from itertools import zip_longest
from typing import override

import xlrd.sheet
from xlrd import Book

from rfexcel.exception.library_exceptions import (LibraryException,
                                                  RowIndexOutOfBoundsException,
                                                  StreamingViolationException)
from rfexcel.utlis.types import Row

from .i_resource import IResource


class XlsEditResource(IResource):
    """Resource for XLS files in standard mode (on_demand=False).
    
    Allows random access to any row. All data is loaded in memory.
    """

    def __init__(self, wb: Book, header_row: int = 1):
        """Initialize XLS Edit Resource.
        
        Arguments:
        - ``wb``: xlrd Book instance
        - ``header_row``: The 1-based row number of the header. (e.g., 1 = Top Row).
        """
        self._wb: Book = wb
        self._active_sheet: xlrd.sheet.Sheet | None = wb.sheet_by_index(0) if wb.nsheets > 0 else None
        self._header_row = header_row
        
        self._headers: list[str] = []
        if self._active_sheet:
            header_idx = header_row - 1
            
            if 0 <= header_idx < self._active_sheet.nrows:
                self._headers = [str(v) for v in self._active_sheet.row_values(header_idx)]

    @override
    def get_row(self, row_index: int) -> Row:
        """Returns row at given index (1-based data index).
        
        Arguments:
        - ``row_index``: 1-based index of the data row.
                        (1 = The first row AFTER the header row)
        """
        if not self._active_sheet:
            raise LibraryException("No active worksheet")
        
        target_xlrd_index = (self._header_row - 1) + row_index
        
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
    """Resource for XLS files in on-demand mode.
    
    Enforces forward-only access based on 1-based data indexing.
    """

    def __init__(self, wb: Book, header_row: int = 1):
        self._wb: Book = wb
        self._active_sheet: xlrd.sheet.Sheet | None = wb.sheet_by_index(0) if wb.nsheets > 0 else None
        self._header_row = header_row
        
        self._last_read_data_index: int = 0
        
        self._headers: list[str] = []
        if self._active_sheet:
            header_idx = header_row - 1
            
            if 0 <= header_idx < self._active_sheet.nrows:
                self._headers = [str(v) for v in self._active_sheet.row_values(header_idx)]

    @override
    def get_row(self, row_index: int) -> Row:
        """Returns row at given index (1-based data index).
        
        Arguments:
        - ``row_index``: 1-based index of the data row.
        """
        if not self._active_sheet:
            raise LibraryException("No active worksheet")
        
        if row_index <= self._last_read_data_index:
            raise StreamingViolationException(row_index, self._last_read_data_index)
        
        self._last_read_data_index = row_index
        
        target_xlrd_index = (self._header_row - 1) + row_index
        
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