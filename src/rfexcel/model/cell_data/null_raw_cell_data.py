from typing import override

from rfexcel.model.cell_data.i_raw_cell_data import IRawCellData
from rfexcel.utils.library_logger import logger


class NullRawCellData(IRawCellData):

    @override
    def get_value(self) -> str:
        logger.warn("No cell data value was returned")
        return ""
