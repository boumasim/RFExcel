from itertools import zip_longest
from typing import Any, override

from openpyxl.cell.cell import Cell
from robot.utils import DotDict  # type: ignore

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utlis.types import DictRowData, ListRowData


class XlsxRawRowData(IRawRowData):
    def __init__(self, data: tuple[Cell, ...] | tuple[Any, ...], value_only: bool):
        self._data = data
        self._value_only = value_only

    @override
    def get_list_row_data(self) -> ListRowData:
        if self._value_only:
            return [str(v) if v is not None else "" for v in self._data]
        return [str(cell.value) if cell.value is not None else "" for cell in self._data]  # type: ignore[union-attr]

    @override
    def get_dict_row_data(self, headers: ListRowData) -> DictRowData:
        if self._value_only:
            values = (str(v) if v is not None else "" for v in self._data)
        else:
            values = (str(cell.value) if cell.value is not None else "" for cell in self._data)  # type: ignore[union-attr]
        return DotDict(zip_longest(headers, values, fillvalue=""))