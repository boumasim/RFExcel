from typing import override

from robot.api import logger

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utlis.types import DictRowData, ListRowData


class NullRawRowData(IRawRowData):

    @override
    def get_list_row_data(self) -> ListRowData:
        logger.warn("No headers were loaded")
        return []
    
    @override
    def get_dict_row_data(self, headers: ListRowData) -> DictRowData:
        logger.warn("No row data values were returned")
        return {}