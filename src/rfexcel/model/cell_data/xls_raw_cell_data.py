from typing import Any, override

from xlrd.sheet import Cell

from rfexcel.model.cell_data.i_raw_cell_data import IRawCellData


class XlsRawCellData(IRawCellData):
    def __init__(self, cell_value: Cell, coordinate: str):
        self._cell_value = cell_value
        self._coordinate = coordinate

    @override
    def get_value(self) -> Any:
        return self._cell_value.value
