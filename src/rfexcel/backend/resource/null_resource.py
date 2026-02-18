from typing import override

from openpyxl.chartsheet import Chartsheet
from openpyxl.worksheet.worksheet import Worksheet

from rfexcel.exception.library_exceptions import LibraryException
from rfexcel.utlis.types import Row

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
    def fetch_row(self, row_index: int) -> Row:
        raise LibraryException("Invalid operation: resource not available")