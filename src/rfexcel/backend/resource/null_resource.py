from typing import override

from openpyxl.chartsheet import Chartsheet
from openpyxl.worksheet.worksheet import Worksheet

from rfexcel.exception.library_exceptions import LibraryException
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData

from .i_resource import IResource


class NullResource(IResource):

    @property
    @override
    def get_active_sheet(self) -> Worksheet | Chartsheet | None:
        raise LibraryException("Invalid operation: resource not available")

    @property
    @override
    def last_read_row_index(self) -> int:
        raise LibraryException("Invalid operation: resource not available")

    @override
    def close(self):
        raise LibraryException("Invalid operation: resource not available")

    @override
    def fetch_row(self, row_index: int, data_only: bool = True) -> IRawRowData:
        raise LibraryException("Invalid operation: resource not available")