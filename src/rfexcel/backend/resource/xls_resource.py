from pathlib import Path
from typing import Any, override

import xlrd.sheet
from xlrd import Book

from rfexcel.exception.library_exceptions import (
    LibraryException, OperationNotSupportedForFormat,
    SheetDoesNotExistException, StreamingViolationException)
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.model.raw_data.xls_raw_row_data import XlsRawRowData
from rfexcel.utils.types import ColumnValues

from .i_resource import IResource


class XlsEditResource(IResource):
    def __init__(self, wb: Book, path: Path):
        super().__init__(path)
        self._wb: Book = wb
        self._active_sheet: xlrd.sheet.Sheet | None = wb.sheet_by_index(0) if wb.nsheets > 0 else None

    @property
    @override
    def active_sheets(self) -> xlrd.sheet.Sheet | None:
        return self._active_sheet
    
    @property
    @override
    def current_sheet(self) -> str:
        if not self._active_sheet:
            raise LibraryException("No active worksheet")
        return self._active_sheet.name

    @property
    @override
    def last_read_row_index(self) -> int:
        return -1

    @override
    def fetch_row(self, row_index: int, **kwargs: Any) -> IRawRowData:
        if not self._active_sheet:
            raise LibraryException("No active worksheet")

        target_xlrd_index = row_index - 1

        if target_xlrd_index >= self._active_sheet.nrows or target_xlrd_index < 0:
            raise StopIteration()

        return XlsRawRowData(list(self._active_sheet.row(target_xlrd_index)))

    @override
    def get_sheet_names(self) -> list[str]:
        return list(self._wb.sheet_names())

    @override
    def switch_sheet(self, name: str) -> None:
        if name not in self._wb.sheet_names():
            raise SheetDoesNotExistException(name)
        self._active_sheet = self._wb.sheet_by_name(name)

    @override
    def add_sheet(self, name: str) -> None:
        raise OperationNotSupportedForFormat(".xls format is read-only; adding sheets is not supported")

    @override
    def delete_sheet(self, name: str) -> None:
        raise OperationNotSupportedForFormat(".xls format is read-only; deleting sheets is not supported")

    @override
    def save(self, path: Path | None = None) -> None:
        raise OperationNotSupportedForFormat(
            ".xls format is read-only; saving is not supported. "
            "Should be converted to .xlsx format before saving, using the XLSWriter utility."
        )

    @override
    def append_row(self, cell_data: ColumnValues) -> None:
        raise OperationNotSupportedForFormat(".xls format is read-only; appending rows is not supported")

    @override
    def update_row(self, row_index: int, cell_data: ColumnValues) -> None:
        raise OperationNotSupportedForFormat(".xls format is read-only; updating rows is not supported")

    @override
    def delete_row(self, row_index: int) -> None:
        raise OperationNotSupportedForFormat(".xls format is read-only; deleting rows is not supported")

    @override
    def insert_row(self, row_index: int, cell_data: ColumnValues) -> None:
        raise OperationNotSupportedForFormat(".xls format is read-only; inserting rows is not supported")

    @override
    def close(self):
        self._wb.release_resources()


class XlsStreamResource(IResource):
    def __init__(self, wb: Book, path: Path):
        super().__init__(path)
        self._wb: Book = wb
        self._active_sheet: xlrd.sheet.Sheet | None = wb.sheet_by_index(0) if wb.nsheets > 0 else None
        self._last_read_row_index = 0

    @property
    @override
    def active_sheets(self) -> xlrd.sheet.Sheet | None:
        return self._active_sheet
    
    @property
    @override
    def current_sheet(self) -> str:
        if not self._active_sheet:
            raise LibraryException("No active worksheet")
        return self._active_sheet.name

    @property
    @override
    def last_read_row_index(self) -> int:
        return self._last_read_row_index

    @override
    def fetch_row(self, row_index: int, **kwargs: Any) -> IRawRowData:
        if not self._active_sheet:
            raise LibraryException("No active worksheet")
        
        if row_index <= self._last_read_row_index:
            raise StreamingViolationException(row_index=row_index, last_read=self._last_read_row_index)

        target_xlrd_index = row_index - 1

        if target_xlrd_index >= self._active_sheet.nrows or target_xlrd_index < 0:
            raise StopIteration()

        self._last_read_row_index = row_index
        return XlsRawRowData(list(self._active_sheet.row(target_xlrd_index)))

    @override
    def get_sheet_names(self) -> list[str]:
        return list(self._wb.sheet_names())

    @override
    def switch_sheet(self, name: str) -> None:
        if name not in self._wb.sheet_names():
            raise SheetDoesNotExistException(name)
        self._active_sheet = self._wb.sheet_by_name(name)
        self._last_read_row_index = 0

    @override
    def add_sheet(self, name: str) -> None:
        raise OperationNotSupportedForFormat(".xls format is read-only; adding sheets is not supported")

    @override
    def delete_sheet(self, name: str) -> None:
        raise OperationNotSupportedForFormat(".xls format is read-only; deleting sheets is not supported")

    @override
    def save(self, path: Path | None = None) -> None:
        raise OperationNotSupportedForFormat(".xls format is read-only; saving is not supported")

    @override
    def append_row(self, cell_data: ColumnValues) -> None:
        raise OperationNotSupportedForFormat(".xls format is read-only; appending rows is not supported")

    @override
    def update_row(self, row_index: int, cell_data: ColumnValues) -> None:
        raise OperationNotSupportedForFormat(".xls format is read-only; updating rows is not supported")

    @override
    def delete_row(self, row_index: int) -> None:
        raise OperationNotSupportedForFormat(".xls format is read-only; deleting rows is not supported")

    @override
    def insert_row(self, row_index: int, cell_data: ColumnValues) -> None:
        raise OperationNotSupportedForFormat(".xls format is read-only; inserting rows is not supported")

    @override
    def close(self):
        self._wb.release_resources()