from abc import ABC, abstractmethod

from rfexcel.utlis.types import DictRowData, HeaderMap, ListRowData


class IRawRowData(ABC):

    @abstractmethod
    def get_list_row_data(self) -> ListRowData:
        pass

    @abstractmethod
    def get_dict_row_data(self, header_map: HeaderMap) -> DictRowData:
        pass

    @abstractmethod
    def get_header_map(self) -> HeaderMap:
        pass