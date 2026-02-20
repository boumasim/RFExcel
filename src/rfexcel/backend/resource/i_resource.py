from abc import ABC, abstractmethod
from typing import Any

from openpyxl.chartsheet import Chartsheet
from openpyxl.worksheet.worksheet import Worksheet

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData


class IResource(ABC):

    @property
    @abstractmethod
    def get_active_sheet(self) -> Any:
        pass

    @property
    @abstractmethod
    def last_read_row_index(self) -> int:
        pass

    @abstractmethod
    def close(self):
        pass

    @abstractmethod
    def fetch_row(self, row_index: int) -> IRawRowData:
        """Return a single row by index (0-based). 
        
        For streaming resources, must be called sequentially. Attempting to read
        a previously read row will raise an exception.
        
        Args:
            row_index: The row index (0-based, data rows only, headers already read)
            
        Returns:
            Dictionary representing the row with headers as keys
            
        Raises:
            IndexError: If row_index is out of bounds
            RuntimeError: If trying to read backwards in streaming mode
        """
        pass
