from typing import override

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.exception.library_exceptions import StreamingViolationException
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.model.raw_data.null_raw_row_data import NullRawRowData

from .i_reader import IReader


class XlsxStreamReader(IReader):

    def __init__(self):
        pass

    @override
    def print(self):
        print("xlsx stream reader\n")

    @override
    def get_headers(self, header_row_idx: int, resource: IResource) -> IRawRowData:
        for i in range(header_row_idx):
            row_data = resource.fetch_row(row_index=i)
            if i == header_row_idx - 1:
                return row_data
        return NullRawRowData()
    
    @override
    def get_row(self, row_idx: int, resource: IResource) -> IRawRowData:
        if row_idx <= resource.last_read_row_index:
            raise StreamingViolationException(row_idx, resource.last_read_row_index)
        
        row_data: IRawRowData = NullRawRowData()
        while resource.last_read_row_index < row_idx:
            row_data = resource.fetch_row(row_index=resource.last_read_row_index)

        return row_data