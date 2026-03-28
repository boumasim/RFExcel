from typing import Any, override

from openpyxl.cell import Cell, MergedCell, ReadOnlyCell
from openpyxl.cell.read_only import EmptyCell

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utils.types import DictRowData, HeaderMap, ListRowData


class XlsxRawRowData(IRawRowData):
    def __init__(self, data: tuple[Cell | MergedCell | ReadOnlyCell | EmptyCell, ...]):
        self._data = data

    @staticmethod
    def _raw_cell_value(cell: Cell | MergedCell | ReadOnlyCell | EmptyCell) -> Any | None:
        if isinstance(cell, EmptyCell):
            return ""
        return cell.value

    @override
    def get_list_row_data(self) -> ListRowData:
        return [
            ("" if (value := self._raw_cell_value(self._data[index])) is None else value)
            for index in range(len(self._data))
        ]

    @override
    def get_dict_row_data(self, header_map: HeaderMap) -> DictRowData:
        row_len = len(self._data)
        result: DictRowData = {}
        for name, col in header_map.items():
            if col <= 0 or col > row_len:
                result[name] = ""
                continue
            value = self._raw_cell_value(self._data[col - 1])
            result[name] = "" if value is None else value
        return result

    @override
    def get_header_map(self) -> HeaderMap:
        return {
            s: i + 1
            for i, cell in enumerate(self._data)
            if (value := self._raw_cell_value(cell)) is not None
            and (s := str(value).strip()) != ""
        }
