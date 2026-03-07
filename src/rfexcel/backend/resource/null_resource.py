from pathlib import Path
from typing import Any, override

from openpyxl.chartsheet import Chartsheet
from openpyxl.worksheet.worksheet import Worksheet

from rfexcel.exception.library_exceptions import LibraryException
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utlis.types import ColumnValues

from .i_resource import IResource


class NullResource(IResource):

    def __init__(self):
        super().__init__(Path(""))

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
    def get_sheet_names(self) -> list[str]:
        raise LibraryException("Invalid operation: resource not available")

    @override
    def switch_sheet(self, name: str) -> None:
        raise LibraryException("Invalid operation: resource not available")

    @override
    def fetch_row(self, row_index: int, **kwargs: Any) -> IRawRowData:
        raise LibraryException("Invalid operation: resource not available")

    @override
    def add_sheet(self, name: str) -> None:
        raise LibraryException("Invalid operation: resource not available")

    @override
    def delete_sheet(self, name: str) -> None:
        raise LibraryException("Invalid operation: resource not available")

    @override
    def save(self, path: Path | None = None) -> None:
        raise LibraryException("Invalid operation: resource not available")

    @override
    def append_row(self, cell_data: ColumnValues) -> None:
        raise LibraryException("Invalid operation: resource not available")