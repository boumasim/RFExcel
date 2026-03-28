from typing import Any, override

import xlrd
from xlrd.sheet import Cell

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utils.types import DictRowData, HeaderMap, ListRowData


class XlsRawRowData(IRawRowData):
    def __init__(self, data: list[Cell]):
        self._data = data

    @staticmethod
    def _norm(cell: Cell) -> Any:
        """Normalize cell value, e.g. convert floats that are actually integers to int."""
        if cell.ctype == xlrd.XL_CELL_NUMBER:
            v: float = float(cell.value)
            if v.is_integer():
                return int(v)
        return cell.value

    @override
    def get_list_row_data(self) -> ListRowData:
        return [self._norm(cell) for cell in self._data]

    @override
    def get_dict_row_data(self, header_map: HeaderMap) -> DictRowData:
        return {
            name: (self._norm(self._data[col - 1]) if 0 < col <= len(self._data) else "")
            for name, col in header_map.items()
        }

    @override
    def get_header_map(self) -> HeaderMap:
        return {
            s: i + 1
            for i, cell in enumerate(self._data)
            if (s := str(self._norm(cell)).strip()) != ""
        }
