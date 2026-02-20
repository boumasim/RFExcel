from abc import ABC, abstractmethod

from rfexcel.utlis.types import Row


class IRawRowData(ABC):

    @abstractmethod
    def get_headers(self) -> list[str]:
        pass

    @abstractmethod
    def get_row_data_value(self, headers: list[str]) -> Row:
        pass