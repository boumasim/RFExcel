from typing import override

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.exception.library_exceptions import LibraryException
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData

from .i_reader import IReader


class NullReader(IReader):

    @override
    def print(self):
        raise LibraryException("Invalid operation: reader not available")
    
    @override
    def get_headers(self, header_row_idx: int, resource: IResource) -> IRawRowData:
        raise LibraryException("Invalid operation: reader not available")
    
    @override
    def get_row(self, row_idx: int, resource: IResource) -> IRawRowData:
        raise LibraryException("Invalid operation: reader not available")