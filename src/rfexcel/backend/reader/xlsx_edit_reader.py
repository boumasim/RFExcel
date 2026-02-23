from typing import Any, override

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData

from .i_reader import IReader


class XlsxEditReader(IReader):

    def __init__(self):
        pass

    @override
    def print(self):
        print("xlsx edit reader\n")

    @override
    def get_headers(self, header_row_idx: int, resource: IResource, **kwargs: Any) -> IRawRowData:
        return resource.fetch_row(row_index=header_row_idx, **kwargs)

    @override
    def get_row(self, row_idx: int, resource: IResource, **kwargs: Any) -> IRawRowData:
        return resource.fetch_row(row_index=row_idx, **kwargs)