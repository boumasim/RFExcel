from typing import override

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utils.types import DictRowData, HeaderMap, ListRowData
from rfexcel.utils.utilities import fast_safe_number_cast


class CsvRawRowData(IRawRowData):
    def __init__(self, data: list[str]):
        self._data = data

    @override
    def get_list_row_data(self) -> ListRowData:
        return list(map(fast_safe_number_cast, self._data))

    @override
    def get_dict_row_data(self, header_map: HeaderMap) -> DictRowData:
        return {
            name: (fast_safe_number_cast(self._data[col - 1]) if 0 < col <= len(self._data) else None)
            for name, col in header_map.items()
        }

    @override
    def get_header_map(self) -> HeaderMap:
        return {
            name.strip(): i + 1
            for i, name in enumerate(self._data)
            if name.strip() != ""
        }