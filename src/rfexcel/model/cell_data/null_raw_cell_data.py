from typing import override

from rfexcel.model.cell_data.i_raw_cell_data import IRawCellData


class NullRawCellData(IRawCellData):
    @override
    def get_value(self) -> str:
        return ""
