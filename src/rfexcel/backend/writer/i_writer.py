from abc import ABC, abstractmethod
from pathlib import Path

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.utils.types import ColumnValues, InsertNativeType


class IWriter(ABC):
    
    @abstractmethod
    def print(self) -> None:
        pass

    @abstractmethod
    def add_sheet(self, name: str, resource: IResource):
        pass

    @abstractmethod
    def delete_sheet(self, name: str, resource: IResource):
        pass

    @abstractmethod
    def save(self, path: Path | None, resource: IResource) -> None:
        pass

    @abstractmethod
    def append_row(self, cell_data: ColumnValues, resource: IResource) -> None:
        pass

    @abstractmethod
    def update_row(self, row_index: int, cell_data: ColumnValues, resource: IResource) -> None:
        pass

    @abstractmethod
    def delete_row(self, row_index: int, resource: IResource) -> None:
        pass

    @abstractmethod
    def insert_row(self, row_index: int, cell_data: ColumnValues, resource: IResource) -> None:
        pass

    @abstractmethod
    def set_cell(self, cell_name: str, value: InsertNativeType, resource: IResource) -> None:
        pass
