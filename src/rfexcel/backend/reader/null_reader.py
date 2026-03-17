from typing import Any, override

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.exception.library_exceptions import NullComponentException
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData

from .i_reader import IReader


class NullReader(IReader):

    @override
    def print(self):
        raise NullComponentException()
    
    @override
    def get_headers(self, header_row_idx: int, resource: IResource, **kwargs: Any) -> IRawRowData:
        raise NullComponentException()
    
    @override
    def get_row(self, row_idx: int, resource: IResource, **kwargs: Any) -> IRawRowData:
        raise NullComponentException()