from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any, Dict, List, Union

from rfexcel.backend.reader.i_reader import IReader
from rfexcel.backend.resource.i_resource import IResource
from rfexcel.backend.writer.i_writer import IWriter
from rfexcel.utils.types import DictRowData, HeaderSpec, ListRowData


class IExcel(ABC):

    @property
    @abstractmethod
    def resource(self) -> IResource:
        pass

    @property
    @abstractmethod
    def reader(self) -> IReader:
        pass
    
    @property
    @abstractmethod
    def writer(self) -> IWriter:
        pass

    @abstractmethod
    def close(self):
        pass

    @abstractmethod
    def get_rows(self,
                header_row: int,
                search_criteria: str | DictRowData | None = None,
                partial_match: bool = False,
                one_row: bool = False,
                **kwargs: Any) -> List[DictRowData] | DictRowData:
        pass

    @abstractmethod
    def list_sheet_names(self) -> list[str]:
        pass

    @abstractmethod
    def switch_sheet(self, name: str) -> None:
        pass

    @abstractmethod
    def add_sheet(self, name: str) -> None:
        pass

    @abstractmethod
    def delete_sheet(self, name: str) -> None:
        pass

    @abstractmethod
    def get_row(self, row: int, headers: HeaderSpec, **kwargs: Any) -> Union[DictRowData, ListRowData]:
        pass

    @abstractmethod
    def save_workbook(self, path: str | None = None) -> None:
        pass

    @abstractmethod
    def append_row(self, row_data: DictRowData, header_row: int) -> None:
        pass

    @abstractmethod
    def append_rows(self, rows: list[DictRowData], header_row: int) -> None:
        pass

    @abstractmethod
    def update_values(self,
                      search_criteria: str | DictRowData,
                      values: str | DictRowData,
                      header_row: int,
                      partial_match: bool,
                      first_only: bool) -> int:
        pass

    @abstractmethod
    def delete_rows(self,
                    search_criteria: str | DictRowData,
                    header_row: int,
                    partial_match: bool,
                    first_only: bool) -> int:
        pass

    @abstractmethod
    def delete_row(self, row_number: int) -> None:
        pass

    @abstractmethod
    def insert_row(self, row_data: DictRowData, row: int, header_row: int) -> None:
        pass

    @abstractmethod
    def compare_data_to(self,
                        target: IExcel,
                        source_header_row: int,
                        target_header_row: int,
                        target_sheet: str | None,
                        headers: list[str] | None,
                        fail_on_diff: bool) -> List[Dict[str, Any]]:
        pass

class ISetExcel(ABC):
    @abstractmethod
    def xls_to_xlsx(self):
        pass