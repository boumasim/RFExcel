from typing import override

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utils.types import DictRowData, HeaderMap, ListRowData


class NullRawRowData(IRawRowData):
    @override
    def get_list_row_data(self) -> ListRowData:
        return []

    @override
    def get_dict_row_data(self, header_map: HeaderMap) -> DictRowData:
        return {}

    @override
    def get_header_map(self) -> HeaderMap:
        return {}
