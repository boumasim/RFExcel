from pathlib import Path
from typing import Any, override

from openpyxl.chartsheet import Chartsheet
from openpyxl.worksheet.worksheet import Worksheet

from rfexcel.exception.library_exceptions import NullComponentException
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utlis.types import ColumnValues

from .i_resource import IResource


class NullResource(IResource):

    def __init__(self):
        super().__init__(Path(""))

    @property
    @override
    def get_active_sheet(self) -> Worksheet | Chartsheet | None:
        raise NullComponentException()

    @property
    @override
    def last_read_row_index(self) -> int:
        raise NullComponentException()

    @override
    def close(self):
        raise NullComponentException()

    @override
    def get_sheet_names(self) -> list[str]:
        raise NullComponentException()

    @override
    def switch_sheet(self, name: str) -> None:
        raise NullComponentException()

    @override
    def fetch_row(self, row_index: int, **kwargs: Any) -> IRawRowData:
        raise NullComponentException()

    @override
    def add_sheet(self, name: str) -> None:
        raise NullComponentException()

    @override
    def delete_sheet(self, name: str) -> None:
        raise NullComponentException()

    @override
    def save(self, path: Path | None = None) -> None:
        raise NullComponentException()

    @override
    def append_row(self, cell_data: ColumnValues) -> None:
        raise NullComponentException()

    @override
    def update_row(self, row_index: int, cell_data: ColumnValues) -> None:
        raise NullComponentException()

    @override
    def delete_row(self, row_index: int) -> None:
        raise NullComponentException()