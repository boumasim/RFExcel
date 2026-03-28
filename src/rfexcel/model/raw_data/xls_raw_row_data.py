from typing import override

from xlrd.sheet import Cell

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utils.types import DictRowData, HeaderMap, ListRowData


class XlsRawRowData(IRawRowData):
    def __init__(self, data: list[Cell]):
        self._data = data

    @override
    def get_list_row_data(self) -> ListRowData:
        return [cell.value for cell in self._data]

    @override
    def get_dict_row_data(self, header_map: HeaderMap) -> DictRowData:
        return {
            name: (self._data[col - 1].value if 0 < col <= len(self._data) else None)
            for name, col in header_map.items()
        }

    @override
    def get_header_map(self) -> HeaderMap:
        return {
            str(cell.value): i + 1
            for i, cell in enumerate(self._data)
            if str(cell.value).strip() != ""
        }
