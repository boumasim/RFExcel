from pathlib import Path
from typing import override

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.exception.library_exceptions import LibraryException
from rfexcel.utlis.types import ColumnValues

from .i_writer import IWriter


class NullWriter(IWriter):

    @override
    def print(self):
        print("will throw writer exception\n")

    @override
    def add_sheet(self, name: str, resource: IResource):
        raise LibraryException("Invalid operation: writer not available")

    @override
    def delete_sheet(self, name: str, resource: IResource):
        raise LibraryException("Invalid operation: writer not available")

    @override
    def save(self, path: Path | None, resource: IResource) -> None:
        raise LibraryException("Invalid operation: writer not available")

    @override
    def add_row(self, cell_data: ColumnValues, resource: IResource) -> None:
        raise LibraryException("Invalid operation: writer not available")