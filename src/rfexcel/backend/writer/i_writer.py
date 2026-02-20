from abc import ABC, abstractmethod

class IWriter(ABC):

    @abstractmethod
    def print(self) -> None:
        pass