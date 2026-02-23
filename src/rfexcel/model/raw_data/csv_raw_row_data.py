from itertools import zip_longest
from typing import override

from robot.utils import DotDict  # type: ignore

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utlis.types import DictRowData, ListRowData


class CsvRawRowData(IRawRowData):
    def __init__(self, data: list[str]):
        self._data = data

    @override
    def get_list_row_data(self) -> ListRowData:
        return list(self._data)

    @override
    def get_dict_row_data(self, headers: ListRowData) -> DictRowData:
        return DotDict(zip_longest(headers, self._data, fillvalue=""))