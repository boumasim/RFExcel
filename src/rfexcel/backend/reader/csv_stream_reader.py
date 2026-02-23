from typing import Any, override

from rfexcel.backend.reader.i_reader import IReader
from rfexcel.backend.resource.i_resource import IResource
from rfexcel.exception.library_exceptions import StreamingViolationException
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.model.raw_data.null_raw_row_data import NullRawRowData


class CsvStreamReader(IReader):
    def __init__(self):
        pass

    @override
    def print(self):
        print("csv stream reader")

    @override
    def get_headers(self, header_row_idx: int, resource: IResource, **kwargs: Any) -> IRawRowData:
        for i in range(header_row_idx):
            row_data = resource.fetch_row(row_index=i, **kwargs)
            if i == header_row_idx - 1:
                return row_data
        return NullRawRowData()

    @override
    def get_row(self, row_idx: int, resource: IResource, **kwargs: Any) -> IRawRowData:
        if row_idx <= resource.last_read_row_index:
            raise StreamingViolationException(row_idx, resource.last_read_row_index)

        row_data: IRawRowData = NullRawRowData()
        while resource.last_read_row_index < row_idx:
            row_data = resource.fetch_row(row_index=resource.last_read_row_index, **kwargs)
        return row_data