from abc import ABC, abstractmethod
from typing import Any


class IRawCellData(ABC):

    @abstractmethod
    def get_value(self) -> Any:
        pass
