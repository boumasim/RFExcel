from abc import ABC, abstractmethod

from rfexcel.backend.resource.i_resource import IResource


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