from pathlib import Path
from typing import Any, List, Union, cast, override

from openpyxl import Workbook
from robot.api import logger
from xls2xlsx import XLS2XLSX

from rfexcel.backend.lib.i_library import IExcel, ISetExcel
from rfexcel.backend.metadata.xlsx_metadata import XlsxMetadata
from rfexcel.backend.reader.xlsx_edit_reader import XlsxEditReader
from rfexcel.backend.resource.xlsx_resource import XlsxEditResource
from rfexcel.backend.style.xlsx_style import XlsxStyle
from rfexcel.backend.writer.xlsx_writer import XlsxWriter
from rfexcel.exception.library_exceptions import HeadersNotDeterminedException
from rfexcel.utlis.utilities import (convert_string_to_dict_row_data,
                                     headers_to_header_map, search_in_row)

from .backend.metadata.i_metadata import IMetadata
from .backend.metadata.null_metadata import NullMetadata
from .backend.reader.i_reader import IReader
from .backend.reader.null_reader import NullReader
from .backend.resource.i_resource import IResource
from .backend.resource.null_resource import NullResource
from .backend.style.i_style import IStyle
from .backend.style.null_style import NullStyle
from .backend.writer.i_writer import IWriter
from .backend.writer.null_writer import NullWriter
from .utlis.types import (ColumnValues, DictRowData, HeaderMap, HeaderSpec,
                          ListRowData, RowInputData)


class RFExcel(IExcel, ISetExcel):

    def __init__(self,
                writer: IWriter = NullWriter(),
                reader: IReader = NullReader(),
                style: IStyle = NullStyle(),
                metadata: IMetadata = NullMetadata(),
                resource: IResource = NullResource()):
        self._writer: IWriter = writer
        self._reader: IReader = reader
        self._style: IStyle = style
        self._metadata: IMetadata = metadata
        self._resource: IResource = resource

    @property
    def writer(self) -> IWriter:
        return self._writer

    @property
    def resource(self) -> IResource:
        return self._resource

    @override
    def close(self):
        self._resource.close()

    @override
    def get_rows(self,
                header_row: int,
                search_criteria: str | RowInputData | None = None,
                partial_match: bool = False,
                one_row: bool = False,
                **kwargs: Any) -> List[DictRowData] | DictRowData:
        search_criteria_dict = convert_string_to_dict_row_data(search_criteria) if search_criteria else None

        try:
            header_map: HeaderMap = self._reader.get_headers(
                header_row_idx=header_row, resource=self._resource, **kwargs
            ).get_header_map()
        except StopIteration:
            header_map = {}

        result: List[DictRowData] = []
        row_index = header_row + 1

        while True:
            try:
                row = self._reader.get_row(row_idx=row_index, resource=self._resource, **kwargs)
                row_dict = row.get_dict_row_data(header_map)
                if not search_criteria_dict or search_in_row(source_row=row_dict, search_criteria=search_criteria_dict, partial_match=partial_match):
                    result.append(row_dict)
                    if one_row:
                        break
                row_index += 1
            except StopIteration:
                break

        return result if not one_row else (result[0] if result else DictRowData())

    @override
    def list_sheet_names(self) -> list[str]:
        return self._metadata.get_sheet_names(self._resource)

    @override
    def switch_sheet(self, name: str) -> None:
        self._resource.switch_sheet(name)

    @override
    def get_row(self, row: int, headers: HeaderSpec, **kwargs: Any) -> Union[DictRowData, ListRowData]:
        try:
            raw = self._reader.get_row(row_idx=row, resource=self._resource, **kwargs)
        except StopIteration:
            return []

        if not headers:
            return raw.get_list_row_data()
        return raw.get_dict_row_data(headers_to_header_map(headers))
    
    @override
    def xls_to_xlsx(self):
        logger.info(
            f"Converting '{self._resource.get_path.name}' from .xls to .xlsx in memory "
            f"to enable write operations. The original .xls file will NOT be modified."
        )
        x2x = XLS2XLSX(str(self._resource.get_path))
        wb : Workbook = cast(Workbook, x2x.to_xlsx()) # type: ignore
        self._resource.close()
        self._resource = XlsxEditResource(wb, self._resource.get_path)
        self._reader = XlsxEditReader()
        self._metadata = XlsxMetadata()
        self._writer = XlsxWriter()
        self._style = XlsxStyle()

    @override
    def add_sheet(self, name: str) -> None:
        self._writer.add_sheet(name=name, resource=self._resource)

    @override
    def delete_sheet(self, name: str):
        self._writer.delete_sheet(name=name, resource=self._resource)

    @override
    def save_workbook(self, path: str | None = None) -> None:
        self._writer.save(Path(path) if path else None, self._resource)

    @override
    def append_row(self, row_data: RowInputData, header_row: int) -> None:
        try:
            header_map = self._reader.get_headers(
                header_row_idx=header_row, resource=self._resource
            ).get_header_map()
        except StopIteration:
            raise HeadersNotDeterminedException(header_row)
        if not header_map:
            raise HeadersNotDeterminedException(header_row)
        cell_data: ColumnValues = {
            col: row_data[name]
            for name, col in header_map.items()
            if name in row_data
        }
        self._writer.append_row(cell_data, self._resource)

    @override
    def append_rows(self, rows: list[RowInputData], header_row: int) -> None:
        for row_data in rows:
            self.append_row(row_data, header_row)

    @override
    def delete_rows(self,
                    search_criteria: str | RowInputData,
                    header_row: int,
                    partial_match: bool,
                    first_only: bool = False) -> int:
        search_criteria_dict = convert_string_to_dict_row_data(search_criteria)
        try:
            header_map: HeaderMap = self._reader.get_headers(
                header_row_idx=header_row, resource=self._resource
            ).get_header_map()
        except StopIteration:
            raise HeadersNotDeterminedException(header_row)
        if not header_map:
            raise HeadersNotDeterminedException(header_row)
        matches: list[int] = []
        row_index = header_row + 1
        while True:
            try:
                row = self._reader.get_row(row_idx=row_index, resource=self._resource)
                row_dict = row.get_dict_row_data(header_map)
                if search_in_row(source_row=row_dict, search_criteria=search_criteria_dict, partial_match=partial_match):
                    matches.append(row_index)
                    if first_only:
                        break
                row_index += 1
            except StopIteration:
                break
        for idx in reversed(matches):
            self._writer.delete_row(idx, self._resource)
        return len(matches)

    @override
    def update_values(self,
                      search_criteria: str | RowInputData,
                      values: str | RowInputData,
                      header_row: int,
                      partial_match: bool,
                      first_only: bool = False) -> int:
        search_criteria_dict = convert_string_to_dict_row_data(search_criteria)
        values_dict = convert_string_to_dict_row_data(values)
        try:
            header_map: HeaderMap = self._reader.get_headers(
                header_row_idx=header_row, resource=self._resource
            ).get_header_map()
        except StopIteration:
            raise HeadersNotDeterminedException(header_row)
        if not header_map:
            raise HeadersNotDeterminedException(header_row)
        update_cell_data: ColumnValues = {
            col: values_dict[name]
            for name, col in header_map.items()
            if name in values_dict
        }
        updated = 0
        row_index = header_row + 1
        while True:
            try:
                row = self._reader.get_row(row_idx=row_index, resource=self._resource)
                row_dict = row.get_dict_row_data(header_map)
                if search_in_row(source_row=row_dict, search_criteria=search_criteria_dict, partial_match=partial_match):
                    self._writer.update_row(row_index, update_cell_data, self._resource)
                    updated += 1
                    if first_only:
                        break
                row_index += 1
            except StopIteration:
                break
        return updated
