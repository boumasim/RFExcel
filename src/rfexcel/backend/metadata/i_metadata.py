from abc import ABC, abstractmethod

from rfexcel.backend.resource.i_resource import IResource


class IMetadata(ABC):

    @abstractmethod
    def print(self) -> None:
        pass

    @abstractmethod
    def get_sheet_names(self, resource: IResource) -> list[str]:
        pass