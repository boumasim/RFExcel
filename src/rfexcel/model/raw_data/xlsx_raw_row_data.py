from typing import override

from openpyxl.cell import MergedCell, ReadOnlyCell, Cell

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utils.types import DictRowData, HeaderMap, ListRowData


class XlsxRawRowData(IRawRowData):
    def __init__(self, data: tuple[Cell | MergedCell | ReadOnlyCell, ...]):
        self._data = data

    @override
    def get_list_row_data(self) -> ListRowData:
        return [cell.value for cell in self._data]

    @override
    def get_dict_row_data(self, header_map: HeaderMap) -> DictRowData:
        col_to_value = {
            cell.column: cell.value
            for cell in self._data
        }
        return {name: col_to_value.get(col) for name, col in header_map.items()}

    @override
    def get_header_map(self) -> HeaderMap:
        return {
            str(cell.value): cell.column
            for cell in self._data
            if cell.value is not None and str(cell.value).strip() != ""
        }
