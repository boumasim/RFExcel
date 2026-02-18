from abc import ABC, abstractmethod

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utlis.types import Data


class IReader(ABC):

    @abstractmethod
    def print(self) -> None:
        pass

    @abstractmethod
    def get_headers(self, header_row_idx: int, resource: IResource) -> IRawRowData:
        pass
    
    @abstractmethod
    def get_row(self, row_idx: int, resource: IResource) -> IRawRowData:
        pass