from typing import override

from rfexcel.backend.reader.i_reader import IReader
from rfexcel.backend.resource.i_resource import IResource
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData


class XlsOnDemandReader(IReader):

    def __init__(self):
        pass

    @override
    def print(self):
        print("xls on_demand reader\n")

    @override
    def get_headers(self, header_row_idx: int, resource: IResource) -> IRawRowData:
        return resource.fetch_row(row_index=header_row_idx)

    @override
    def get_row(self, row_idx: int, resource: IResource) -> IRawRowData:
        return resource.fetch_row(row_index=row_idx)