from itertools import zip_longest
from typing import override, Optional, Any, List

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from rfexcel.exception.library_exceptions import (
    LibraryException,
    RowIndexOutOfBoundsException,
    StreamingViolationException
)
from rfexcel.utlis.types import Row
from .i_resource import IResource


class XlsxEditResource(IResource):
    """Resource for XLSX files in edit mode (read_only=False).
    
    Allows random access to any row. All data is loaded in memory.
    """

    def __init__(self, wb: Workbook, header_row: int = 1):
        """Initialize XLSX Edit Resource.
        
        Arguments:
        - ``wb``: openpyxl Workbook instance (read_only=False)
        - ``header_row``: Row number where headers are located (1-based). Defaults to 1.
        """
        self._wb: Workbook = wb
        self._active_sheet: Worksheet | None = wb.active
        self._header_row = header_row
        
        self._headers: list[str] = []
        
        if self._active_sheet:
            header_generator = self._active_sheet.iter_rows(
                min_row=header_row,
                max_row=header_row,
                values_only=True
            )
            
            header_data = next(header_generator, None)
            
            if header_data:
                self._headers = [str(cell) if cell is not None else "" for cell in header_data]

    @override
    def get_row(self, row_index: int) -> Row:
        """Returns row at given index (1-based data index).
        
        Supports random access to any row using direct cell access.
        
        Arguments:
        - ``row_index``: 1-based index of data row (row 1 is first data row after headers)
        """
        if not self._active_sheet:
            raise LibraryException("No active worksheet")
        
        target_excel_row = self._header_row + row_index
        
        if target_excel_row > self._active_sheet.max_row:
            raise RowIndexOutOfBoundsException(row_index, f"Data row {row_index} (Excel row {target_excel_row}) out of range")
        
        row_values = next(
            self._active_sheet.iter_rows(min_row=target_excel_row, max_row=target_excel_row, values_only=True),
            None
        )
        
        if row_values is None:
            raise RowIndexOutOfBoundsException(row_index)

        return {
            header: str(val) if val is not None else ""
            for header, val in zip_longest(self._headers, row_values, fillvalue=None)
            if header is not None
        }

    @override
    def close(self):
        """Close the workbook and release resources."""
        self._wb.close()


class XlsxStreamResource(IResource):
    """Resource for XLSX files in streaming mode (read_only=True).
    
    Memory efficient for large files. Only supports forward-only sequential access.
    """

    def __init__(self, wb: Workbook, header_row: int = 1):
        """Initialize XLSX Stream Resource.
        
        Arguments:
        - ``wb``: openpyxl Workbook instance (read_only=True)
        - ``header_row``: Row number where headers are located (1-based). Defaults to 1.
        """
        self._wb = wb
        self._active_sheet = wb.active
        self._headers = []
        self._header_row = header_row
        
        self._last_read_data_index = 0

        if self._active_sheet:
            self._row_generator = self._active_sheet.iter_rows(values_only=True)
            
            for i in range(header_row):
                try:
                    row_data = next(self._row_generator)
                    if i == header_row - 1:
                        self._headers = [str(cell) if cell is not None else "" for cell in row_data]
                except StopIteration:
                    self._headers = []
                    break
        else:
            self._row_generator = iter([])

    @override
    def get_row(self, row_index: int) -> Row:
        """Returns row at given index (1-based, data rows only).
        
        Arguments:
        - ``row_index``: 1-based index of data row.
        """
        
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
        """Close the workbook and release resources."""
        self._wb.close()