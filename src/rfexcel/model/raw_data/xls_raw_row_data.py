from typing import override

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utlis.types import DictRowData, HeaderMap, ListRowData


class XlsRawRowData(IRawRowData):
    def __init__(self, data: list[str | int | float | bool]):
        self._data = data

    @override
    def get_list_row_data(self) -> ListRowData:
        return [str(v) for v in self._data]

    @override
    def get_dict_row_data(self, header_map: HeaderMap) -> DictRowData:
        return DictRowData({
            name: (str(self._data[col - 1]) if col - 1 < len(self._data) else "")
            for name, col in header_map.items()
        })

    @override
    def get_header_map(self) -> HeaderMap:
        return {
            str(v): i + 1
            for i, v in enumerate(self._data)
            if str(v).strip() != ""
        }