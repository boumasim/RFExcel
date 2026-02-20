from itertools import zip_longest
from typing import override

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utlis.types import Row


class CsvRawRowData(IRawRowData):
    def __init__(self, data: list[str]):
        self._data = data

    @override
    def get_headers(self) -> list[str]:
        return [v if v is not None else "" for v in self._data]

    @override
    def get_row_data_value(self, headers: list[str]) -> Row:
        return dict(zip_longest(headers, (v if v is not None else "" for v in self._data), fillvalue=""))