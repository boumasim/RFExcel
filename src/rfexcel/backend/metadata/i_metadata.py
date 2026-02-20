from abc import ABC, abstractmethod


class IMetadata(ABC):

    @abstractmethod
    def print(self) -> None:
        pass