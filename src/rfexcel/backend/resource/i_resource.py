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
    def fetch_row(self, row_index: int, data_only: bool = True) -> IRawRowData:
        """Return a single row by index (1-based).

        Args:
            row_index: The row index (1-based, matches Excel row numbering).
            data_only: When ``True`` (default), returns raw Python values.
                       When ``False``, returns native cell objects (e.g. openpyxl
                       ``Cell``) preserving formula and style metadata.
                       Has no effect for formats that do not support formulas
                       (xls via xlrd, csv).

        Raises:
            StopIteration: If row_index is out of bounds.
            StreamingViolationException: If trying to read backwards in streaming mode.
        """
        pass
