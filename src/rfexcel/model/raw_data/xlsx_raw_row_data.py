from typing import Any, override

from itertools import zip_longest

from openpyxl.cell.cell import Cell

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utlis.types import Row


class XlsxRawRowData(IRawRowData):
    def __init__(self, data: tuple[Cell, ...] | tuple[Any, ...], value_only: bool):
        self._data = data
        self._value_only = value_only

    @override
    def get_headers(self) -> list[str]:
        return [str(cell) if cell is not None else "" for cell in self._data]
    
    @override
    def get_row_data_value(self, headers: list[str]) -> Row:
        return dict(zip_longest(headers, (str(v) if v is not None else "" for v in self._data), fillvalue=""))