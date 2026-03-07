from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any

from rfexcel.model.raw_data.i_raw_row_data import IRawRowData


class IResource(ABC):

    def __init__(self, path: Path):
        self._path: Path = path

    @property
    @abstractmethod
    def get_active_sheet(self) -> Any:
        pass

    @property
    @abstractmethod
    def last_read_row_index(self) -> int:
        pass

    @property
    def get_path(self) -> Path:
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
    def add_sheet(self, name: str) -> None:
        pass

    @abstractmethod
    def delete_sheet(self, name: str) -> None:
        pass

    @abstractmethod
    def save(self, path: Path | None = None) -> None:
        pass
