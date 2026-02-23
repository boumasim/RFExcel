from abc import ABC, abstractmethod

from rfexcel.utlis.types import DictRowData, ListRowData


class IRawRowData(ABC):

    @abstractmethod
    def get_list_row_data(self) -> ListRowData:
        pass

    @abstractmethod
    def get_dict_row_data(self, headers: ListRowData) -> DictRowData:
        pass