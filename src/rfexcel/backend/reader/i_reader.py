from abc import ABC, abstractmethod

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.utlis.types import Data


class IReader(ABC):

    @abstractmethod
    def print(self) -> None:
        pass

    @abstractmethod
    def get_rows(self, resource: IResource) -> Data:
        """Read all rows from resource."""
        pass