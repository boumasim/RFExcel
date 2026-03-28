from typing import override

from openpyxl.cell import Cell, MergedCell, ReadOnlyCell
from openpyxl.cell.read_only import EmptyCell

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utils.types import DictRowData, HeaderMap, ListRowData


class XlsxRawRowData(IRawRowData):
    def __init__(self, data: tuple[Cell | MergedCell | ReadOnlyCell | EmptyCell, ...]):
        self._data = data

    @override
    def get_list_row_data(self) -> ListRowData:
        return [cell.value for cell in self._data]

    @override
    def get_dict_row_data(self, header_map: HeaderMap) -> DictRowData:
        wanted: dict[int, str] = {col: name for name, col in header_map.items()}
        result: DictRowData = {name: None for name in header_map}
        for cell in self._data:
            if not isinstance(cell, EmptyCell) and cell.column in wanted:
                result[wanted[cell.column]] = cell.value
        return result

    @override
    def get_header_map(self) -> HeaderMap:
        return {
            s: cell.column
            for cell in self._data
            if not isinstance(cell, EmptyCell)
            and cell.value is not None
            and (s := str(cell.value)).strip() != ""
        }
