from typing import Any, override

from openpyxl.cell import Cell, MergedCell, ReadOnlyCell

from rfexcel.model.cell_data.i_raw_cell_data import IRawCellData


class XlsxRawCellData(IRawCellData):
    def __init__(self, cell_value: Cell | ReadOnlyCell | MergedCell, coordinate: str):
        self._cell_value = cell_value
        self._coordinate = coordinate

    @override
    def get_value(self) -> Any:
        return self._cell_value.value
