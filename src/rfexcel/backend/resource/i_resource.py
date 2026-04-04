from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any

from rfexcel.model.cell_data.i_raw_cell_data import IRawCellData
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utils.types import ColumnValues, InsertNativeType


class IResource(ABC):

    def __init__(self, path: Path):
        self._path: Path = path

    @property
    @abstractmethod
    def active_sheets(self) -> Any:
        pass

    @property
    @abstractmethod
    def current_sheet(self) -> str:
        pass

    @property
    @abstractmethod
    def last_read_row_index(self) -> int:
        pass

    @property
    def path(self) -> Path:
        return self._path

    @abstractmethod
    def close(self):
        pass

    @abstractmethod
    def get_sheet_names(self) -> list[str]:
        pass

    @abstractmethod
    def switch_sheet(self, name: str) -> None:
        pass

    @abstractmethod
    def fetch_row(self, row_index: int, **kwargs: Any) -> IRawRowData:
        pass

    @abstractmethod
    def fetch_cell(self, cell_name: str, **kwargs: Any) -> IRawCellData:
        pass

    @abstractmethod
    def add_sheet(self, name: str) -> None:
        pass

    @abstractmethod
    def delete_sheet(self, name: str) -> None:
        pass

    @abstractmethod
    def save(self, path: Path | None = None) -> None:
        pass

    @abstractmethod
    def append_row(self, cell_data: ColumnValues) -> None:
        pass

    @abstractmethod
    def update_row(self, row_index: int, cell_data: ColumnValues) -> None:
        pass

    @abstractmethod
    def delete_row(self, row_index: int) -> None:
        pass

    @abstractmethod
    def insert_row(self, row_index: int, cell_data: ColumnValues) -> None:
        pass

    @abstractmethod
    def set_cell(self, cell_name: str, value: InsertNativeType) -> None:
        pass
