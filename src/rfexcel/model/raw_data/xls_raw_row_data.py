from typing import override

from xlrd.sheet import Cell

from rfexcel.model.common_model import norm_xls_value
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utils.types import DictRowData, HeaderMap, ListRowData


class XlsRawRowData(IRawRowData):
    def __init__(self, data: list[Cell]):
        self._data = data

    @override
    def get_list_row_data(self) -> ListRowData:
        return [
            value
            for cell in self._data
            if (value := norm_xls_value(cell)) not in (None, "")
        ]

    @override
    def get_dict_row_data(self, header_map: HeaderMap) -> DictRowData:
        return {
            name: (norm_xls_value(self._data[col - 1]) if 0 < col <= len(self._data) else "")
            for name, col in header_map.items()
        }

    @override
    def get_header_map(self) -> HeaderMap:
        return {
            s: i + 1
            for i, cell in enumerate(self._data)
            if (s := str(norm_xls_value(cell)).strip()) != ""
        }
