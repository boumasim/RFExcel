from typing import Any, override

from rfexcel.backend.reader.i_reader import IReader
from rfexcel.backend.resource.i_resource import IResource
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.model.raw_data.null_raw_row_data import NullRawRowData


class XlsOnDemandReader(IReader):

    def __init__(self):
        pass

    @override
    def print(self):
        print("xls on_demand reader\n")

    @override
    def get_headers(self, header_row_idx: int, resource: IResource, **kwargs: Any) -> IRawRowData:
        return resource.fetch_row(row_index=header_row_idx, **kwargs)

    @override
    def get_row(self, row_idx: int, resource: IResource, **kwargs: Any) -> IRawRowData:
        if row_idx <= 0:
            return NullRawRowData()

        return resource.fetch_row(row_index=row_idx, **kwargs)