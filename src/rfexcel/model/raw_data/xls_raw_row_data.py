from itertools import zip_longest
from typing import override

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utlis.types import DictRowData, ListRowData


class XlsRawRowData(IRawRowData):
    def __init__(self, data: list[str | int | float | bool]):
        self._data = data

    @override
    def get_list_row_data(self) -> ListRowData:
        return [str(v) for v in self._data]

    @override
    def get_dict_row_data(self, headers: ListRowData) -> DictRowData:
        return dict(zip_longest(headers, (str(v) for v in self._data), fillvalue=""))