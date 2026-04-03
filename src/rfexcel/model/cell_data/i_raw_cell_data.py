from abc import ABC, abstractmethod

from rfexcel.utils.types import NativeType


class IRawCellData(ABC):

    @abstractmethod
    def get_value(self) -> NativeType:
        pass
