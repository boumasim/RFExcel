from abc import ABC, abstractmethod


class IStyle(ABC):

    @abstractmethod
    def print(self) -> None:
        pass