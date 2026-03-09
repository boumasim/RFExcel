from typing import Any, override

from openpyxl.cell.cell import Cell

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utlis.types import (ColumnValues, DictRowData, HeaderMap,
                                 ListRowData)


class XlsxRawRowData(IRawRowData):
    def __init__(self, data: tuple[Cell, ...] | tuple[Any, ...], value_only: bool):
        self._data = data
        self._value_only = value_only

    @override
    def get_list_row_data(self) -> ListRowData:
        if self._value_only:
            return [str(v) if v is not None else "" for v in self._data]
        return [str(cell.value) if cell.value is not None else "" for cell in self._data]  # type: ignore[union-attr]

    @override
    def get_dict_row_data(self, header_map: HeaderMap) -> DictRowData:
        if self._value_only:
            return DictRowData({
                name: (
                    str(self._data[col - 1])
                    if col - 1 < len(self._data) and self._data[col - 1] is not None
                    else ""
                )
                for name, col in header_map.items()
            })
        col_to_value: ColumnValues = {
            cell.column: (str(cell.value) if cell.value is not None else "")  # type: ignore[union-attr]
            for cell in self._data
        }
        return DictRowData({name: col_to_value.get(col, "") for name, col in header_map.items()})

    @override
    def get_header_map(self) -> HeaderMap:
        if self._value_only:
            return {
                str(v): i + 1
                for i, v in enumerate(self._data)
                if str(v).strip() != ""
            }
        return {
            str(cell.value): cell.column  # type: ignore[union-attr]
            for cell in self._data
            if cell.value is not None and str(cell.value).strip() != ""  # type: ignore[union-attr]
        }