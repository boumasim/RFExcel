from typing import cast, override

from openpyxl.cell import Cell, MergedCell, ReadOnlyCell
from openpyxl.cell.read_only import EmptyCell

from rfexcel.model.cell_data.i_raw_cell_data import IRawCellData
from rfexcel.utils.types import NativeType


class XlsxRawCellData(IRawCellData):
    def __init__(self, cell_value: Cell | ReadOnlyCell | MergedCell | EmptyCell, coordinate: str):
        self._cell_value = cell_value
        self._coordinate = coordinate

    @override
    def get_value(self) -> NativeType:
        if isinstance(self._cell_value, EmptyCell):
            return ""
        return "" if self._cell_value.value is None else cast(NativeType, self._cell_value.value)
