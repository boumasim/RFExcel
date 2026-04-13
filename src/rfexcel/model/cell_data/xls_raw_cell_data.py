from typing import override

from xlrd.sheet import Cell

from rfexcel.model.cell_data.i_raw_cell_data import IRawCellData
from rfexcel.model.common_model import norm_xls_value
from rfexcel.utils.types import NativeType


class XlsRawCellData(IRawCellData):
    def __init__(self, cell_value: Cell, coordinate: str):
        self._cell_value = cell_value
        self._coordinate = coordinate

    @override
    def get_value(self) -> NativeType:
        return norm_xls_value(self._cell_value)
