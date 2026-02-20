from itertools import zip_longest
from typing import Any, override

from openpyxl.cell.cell import Cell

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utlis.types import Row


class XlsxRawRowData(IRawRowData):
    def __init__(self, data: tuple[Cell, ...] | tuple[Any, ...], value_only: bool):
        self._data = data
        self._value_only = value_only

    @override
    def get_headers(self) -> list[str]:
        if self._value_only:
            return [str(v) if v is not None else "" for v in self._data]
        return [str(cell.value) if cell.value is not None else "" for cell in self._data]  # type: ignore[union-attr]

    @override
    def get_row_data_value(self, headers: list[str]) -> Row:
        if self._value_only:
            values = (str(v) if v is not None else "" for v in self._data)
        else:
            values = (str(cell.value) if cell.value is not None else "" for cell in self._data)  # type: ignore[union-attr]
        return dict(zip_longest(headers, values, fillvalue=""))