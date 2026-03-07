from pathlib import Path
from typing import override

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.backend.writer.i_writer import IWriter
from rfexcel.exception.library_exceptions import OperationNotSupportedForFormat


class CsvWriter(IWriter):
    def __init__(self):
        pass

    @override
    def print(self):
        print("csv reader")

    @override
    def add_sheet(self, name: str, resource: IResource):
        raise OperationNotSupportedForFormat("Adding sheets is not supported for CSV format")

    @override
    def delete_sheet(self, name: str, resource: IResource):
        raise OperationNotSupportedForFormat("Deleting sheets is not supported for CSV format")

    @override
    def save(self, path: Path | None, resource: IResource) -> None:
        resource.save(path)