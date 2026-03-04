from typing import override

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.exception.library_exceptions import LibraryException

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