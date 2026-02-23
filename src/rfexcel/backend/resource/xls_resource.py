from typing import Any, override

import xlrd.sheet
from xlrd import Book

from rfexcel.exception.library_exceptions import LibraryException
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.model.raw_data.xls_raw_row_data import XlsRawRowData

from .i_resource import IResource


class XlsEditResource(IResource):
    def __init__(self, wb: Book):
        self._wb: Book = wb
        self._active_sheet: xlrd.sheet.Sheet | None = wb.sheet_by_index(0) if wb.nsheets > 0 else None

    @property
    @override
    def get_active_sheet(self) -> xlrd.sheet.Sheet | None:
        return self._active_sheet

    @property
    @override
    def last_read_row_index(self) -> int:
        return -1

    @override
    def fetch_row(self, row_index: int, **kwargs: Any) -> IRawRowData:
        if not self._active_sheet:
            raise LibraryException("No active worksheet")

        target_xlrd_index = row_index - 1

        if target_xlrd_index >= self._active_sheet.nrows:
            raise StopIteration()

        return XlsRawRowData(list(self._active_sheet.row_values(target_xlrd_index)))

    @override
    def close(self):
        self._wb.release_resources()


class XlsStreamResource(IResource):
    def __init__(self, wb: Book):
        self._wb: Book = wb
        self._active_sheet: xlrd.sheet.Sheet | None = wb.sheet_by_index(0) if wb.nsheets > 0 else None

    @property
    @override
    def get_active_sheet(self) -> xlrd.sheet.Sheet | None:
        return self._active_sheet

    @property
    @override
    def last_read_row_index(self) -> int:
        return -1

    @override
    def fetch_row(self, row_index: int, **kwargs: Any) -> IRawRowData:
        if not self._active_sheet:
            raise LibraryException("No active worksheet")

        target_xlrd_index = row_index - 1

        if target_xlrd_index >= self._active_sheet.nrows:
            raise StopIteration()

        return XlsRawRowData(list(self._active_sheet.row_values(target_xlrd_index)))

    @override
    def close(self):
        self._wb.release_resources()