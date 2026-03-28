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
        if resource.last_read_row_index >= header_row_idx:
            raise StreamingViolationException(header_row_idx, resource.last_read_row_index)
        try:
            row_data = resource.fetch_row(row_index=header_row_idx, **kwargs)
            return row_data
        except StopIteration:
            return NullRawRowData()

    @override
    def get_row(self, row_idx: int, resource: IResource, **kwargs: Any) -> IRawRowData:
        if row_idx <= resource.last_read_row_index:
            raise StreamingViolationException(row_idx, resource.last_read_row_index)

        return resource.fetch_row(row_index=row_idx, **kwargs)