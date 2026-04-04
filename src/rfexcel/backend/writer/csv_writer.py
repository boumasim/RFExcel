from pathlib import Path
from typing import override

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.backend.writer.i_writer import IWriter
from rfexcel.exception.library_exceptions import OperationNotSupportedForFormat
from rfexcel.utils.types import ColumnValues, InsertNativeType


class CsvWriter(IWriter):
    def __init__(self):
        pass

    @override
    def print(self):
        print("csv writer")

    @override
    def add_sheet(self, name: str, resource: IResource):
        raise OperationNotSupportedForFormat("Adding sheets is not supported for CSV format")

    @override
    def delete_sheet(self, name: str, resource: IResource):
        raise OperationNotSupportedForFormat("Deleting sheets is not supported for CSV format")

    @override
    def save(self, path: Path | None, resource: IResource) -> None:
        resource.save(path)

    @override
    def append_row(self, cell_data: ColumnValues, resource: IResource) -> None:
        resource.append_row(cell_data)

    @override
    def update_row(self, row_index: int, cell_data: ColumnValues, resource: IResource) -> None:
        resource.update_row(row_index, cell_data)

    @override
    def delete_row(self, row_index: int, resource: IResource) -> None:
        resource.delete_row(row_index)

    @override
    def insert_row(self, row_index: int, cell_data: ColumnValues, resource: IResource) -> None:
        resource.insert_row(row_index, cell_data)

    @override
    def set_cell(self, cell_name: str, value: InsertNativeType, resource: IResource) -> None:
        raise OperationNotSupportedForFormat("Set Cell is supported only for .xlsx and .xls files")