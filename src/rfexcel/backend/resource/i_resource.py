from abc import ABC, abstractmethod
from typing import Any

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
    def get_sheet_names(self) -> list[str]:
        """Return the list of sheet names in the workbook."""
        pass

    @abstractmethod
    def switch_sheet(self, name: str) -> None:
        """Switch the active sheet to the one identified by ``name``."""
        pass

    @abstractmethod
    def fetch_row(self, row_index: int, **kwargs: Any) -> IRawRowData:
        """Return a single row by index (1-based).

        Args:
            row_index: The row index (1-based, matches Excel row numbering).
            **kwargs:  Backend-specific options forwarded from the keyword layer
                       (e.g. ``data_only=True`` for openpyxl).
                       ``Cell``) preserving formula and style metadata.
                       Has no effect for formats that do not support formulas
                       (xls via xlrd, csv).

        Raises:
            StopIteration: If row_index is out of bounds.
            StreamingViolationException: If trying to read backwards in streaming mode.
        """
        pass
