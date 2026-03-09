from pathlib import Path
from typing import override

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.utlis.types import ColumnValues

from .i_writer import IWriter


class XlsxWriter(IWriter):

    def __init__(self):
        pass

    @override
    def print(self):
        print("xlsx writer\n")

    @override
    def add_sheet(self, name: str, resource: IResource):
        resource.add_sheet(name)

    @override
    def delete_sheet(self, name: str, resource: IResource):
        resource.delete_sheet(name)

    @override
    def save(self, path: Path | None, resource: IResource) -> None:
        resource.save(path)

    @override
    def add_row(self, cell_data: ColumnValues, resource: IResource) -> None:
        resource.append_row(cell_data)

    @override
    def update_row(self, row_index: int, cell_data: ColumnValues, resource: IResource) -> None:
        resource.update_row(row_index, cell_data)

    @override
    def delete_row(self, row_index: int, resource: IResource) -> None:
        resource.delete_row(row_index)