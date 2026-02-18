from typing import override
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utlis.types import Row
from robot.api import logger


class NullRawRowData(IRawRowData):

    @override
    def get_headers(self) -> list[str]:
        logger.warn("No headers were loaded")
        return []
    
    @override
    def get_row_data_value(self, headers: list[str]) -> Row:
        logger.warn("No row data values were returned")
        return {}