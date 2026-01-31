from abc import ABC, abstractmethod


class IReader(ABC):

    @abstractmethod
    def print(self) -> None:
        pass