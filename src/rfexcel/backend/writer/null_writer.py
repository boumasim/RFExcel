from pathlib import Path
from typing import override

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.exception.library_exceptions import NullComponentException
from rfexcel.utlis.types import ColumnValues

from .i_writer import IWriter


class NullWriter(IWriter):

    @override
    def print(self):
        print("will throw writer exception\n")

    @override
    def add_sheet(self, name: str, resource: IResource):
        raise NullComponentException()

    @override
    def delete_sheet(self, name: str, resource: IResource):
        raise NullComponentException()

    @override
    def save(self, path: Path | None, resource: IResource) -> None:
        raise NullComponentException()

    @override
    def append_row(self, cell_data: ColumnValues, resource: IResource) -> None:
        raise NullComponentException()

    @override
    def update_row(self, row_index: int, cell_data: ColumnValues, resource: IResource) -> None:
        raise NullComponentException()

    @override
    def delete_row(self, row_index: int, resource: IResource) -> None:
        raise NullComponentException()