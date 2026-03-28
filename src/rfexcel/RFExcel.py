from pathlib import Path
from typing import Any, List, Union, override

from openpyxl import Workbook

from rfexcel.backend.metadata.xlsx_metadata import XlsxMetadata
from rfexcel.backend.reader.xlsx_edit_reader import XlsxEditReader
from rfexcel.backend.resource.xlsx_resource import XlsxEditResource
from rfexcel.backend.style.xlsx_style import XlsxStyle
from rfexcel.backend.writer.xlsx_writer import XlsxWriter
from rfexcel.exception.library_exceptions import (
    HeadersNotDeterminedException, NotMatchingColumns,
    RowIndexOutOfBoundsException)
from rfexcel.utils.library_logger import logger
from rfexcel.utils.utilities import (convert_string_to_dict_row_data,
                                     convert_xls_to_xlsx,
                                     headers_to_header_map, search_in_row)

from .backend.interfaces.i_library import IExcel, ISetExcel
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
from .utils.types import (ColumnDifference, ColumnValues, DictRowData,
                          HeaderMap, HeaderSpec, ListRowData, RowDifference)


class RFExcel(IExcel, ISetExcel):

    def __init__(self,
                read_only: bool,
                writer: IWriter = NullWriter(),
                reader: IReader = NullReader(),
                style: IStyle = NullStyle(),
                metadata: IMetadata = NullMetadata(),
                resource: IResource = NullResource()):
        self._read_only = read_only
        self._writer: IWriter = writer
        self._reader: IReader = reader
        self._style: IStyle = style
        self._metadata: IMetadata = metadata
        self._resource: IResource = resource

    @property
    @override
    def read_only(self) -> bool:
        return self._read_only

    @property
    @override
    def writer(self) -> IWriter:
        return self._writer

    @property
    @override
    def resource(self) -> IResource:
        return self._resource
    
    @property
    @override
    def reader(self) -> IReader:
        return self._reader

    @staticmethod
    def _read_header_map(reader: IReader, resource: IResource, header_row: int, **kwargs: Any) -> HeaderMap:
        try:
            return reader.get_headers(header_row_idx=header_row, resource=resource, **kwargs).get_header_map()
        except StopIteration:
            raise HeadersNotDeterminedException(header_row)

    @override
    def close(self):
        self._resource.close()

    @override
    def get_rows(self,
                header_row: int,
                search_criteria: DictRowData | str | None = None,
                partial_match: bool = False,
                one_row: bool = False,
                **kwargs: Any) -> List[DictRowData] | DictRowData:
        search_criteria_dict = convert_string_to_dict_row_data(search_criteria) if search_criteria is not None else None

        header_map: HeaderMap = self._read_header_map(self._reader, self._resource, header_row, **kwargs)

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

        return result if not one_row else (result[0] if result else {})

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
            f"Converting '{self._resource.path.name}' from .xls to .xlsx in memory "
            f"to enable write operations. The original .xls file will NOT be modified."
        )
        wb: Workbook = convert_xls_to_xlsx(Path(self._resource.path))
        new_path: Path = self._resource.path.with_suffix('.xlsx')
        self._resource.close()
        self._resource = XlsxEditResource(wb, new_path)
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
    def append_row(self, row_data: DictRowData, header_row: int) -> None:
        header_map: HeaderMap = self._read_header_map(self._reader, self._resource, header_row)
        if not header_map:
            raise HeadersNotDeterminedException(header_row)
        cell_data: ColumnValues = {
            col: row_data[name]
            for name, col in header_map.items()
            if name in row_data
        }
        self._writer.append_row(cell_data, self._resource)

    @override
    def append_rows(self, rows: list[DictRowData], header_row: int) -> None:
        for row_data in rows:
            self.append_row(row_data, header_row)

    @override
    def insert_row(self, row_data: DictRowData, row: int, header_row: int) -> None:
        if row <= header_row:
            raise RowIndexOutOfBoundsException(
                row, f"Row {row} must be greater than header_row {header_row}"
            )
        header_map: HeaderMap = self._read_header_map(self._reader, self._resource, header_row)
        if not header_map:
            raise HeadersNotDeterminedException(header_row)
        cell_data: ColumnValues = {
            col: row_data[name]
            for name, col in header_map.items()
            if name in row_data
        }
        self._writer.insert_row(row, cell_data, self._resource)

    @override
    def delete_rows(self,
                    search_criteria: DictRowData | str,
                    header_row: int,
                    partial_match: bool,
                    first_only: bool = False) -> int:
        search_criteria_dict = convert_string_to_dict_row_data(search_criteria)
        header_map: HeaderMap = self._read_header_map(self._reader, self._resource, header_row)
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
    def delete_row(self, row_number: int) -> None:
        if row_number < 1:
            raise RowIndexOutOfBoundsException(row_number)
        try:
            self._reader.get_row(row_idx=row_number, resource=self._resource)
        except StopIteration:
            raise RowIndexOutOfBoundsException(row_number)
        self._writer.delete_row(row_number, self._resource)

    @override
    def update_values(self,
                      search_criteria: DictRowData | str,
                      values: str | DictRowData,
                      header_row: int,
                      partial_match: bool,
                      first_only: bool = False) -> int:
        search_criteria_dict = convert_string_to_dict_row_data(search_criteria)
        values_dict = convert_string_to_dict_row_data(values)
        header_map: HeaderMap = self._read_header_map(self._reader, self._resource, header_row)
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

    @override
    def compare_data_to(self,
                        target: IExcel,
                        source_header_row: int,
                        target_header_row: int,
                        target_sheet: str | None,
                        headers: list[str] | None,
                        fail_on_diff: bool) -> list[RowDifference]:
        try:
            if target_sheet is not None:
                target.switch_sheet(target_sheet)

            source_header_map: HeaderMap = self._read_header_map(self._reader, self._resource, source_header_row)
            target_header_map: HeaderMap = self._read_header_map(target.reader, target.resource, target_header_row)

            if headers is None:
                compare_headers: list[str] = list(source_header_map.keys())
                missing_in_target = [h for h in compare_headers if h not in target_header_map]
                if missing_in_target:
                    raise NotMatchingColumns(missing_in_source=[], missing_in_target=missing_in_target)
            else:
                compare_headers = headers
                missing_in_source = [h for h in compare_headers if h not in source_header_map]
                missing_in_target = [h for h in compare_headers if h not in target_header_map]
                if missing_in_source or missing_in_target:
                    raise NotMatchingColumns(missing_in_source=missing_in_source, missing_in_target=missing_in_target)

            result: list[RowDifference] = []
            source_row_index = source_header_row + 1
            target_row_index = target_header_row + 1
            target_exhausted = False

            while True:
                try:
                    source_row = self._reader.get_row(row_idx=source_row_index, resource=self._resource)
                except StopIteration:
                    break

                source_dict = source_row.get_dict_row_data(source_header_map)

                if not target_exhausted:
                    try:
                        target_row = target.reader.get_row(row_idx=target_row_index, resource=target.resource)
                        target_dict: DictRowData = target_row.get_dict_row_data(target_header_map)
                        target_row_index += 1
                    except StopIteration:
                        target_exhausted = True
                        target_dict = {}
                else:
                    target_dict = {}

                source_values = source_dict
                target_values = target_dict
                differences: ColumnDifference = {
                    h: {"source": source_values.get(h), "target": target_values.get(h)}
                    for h in compare_headers
                    if source_values.get(h) != target_values.get(h)
                }

                if differences:
                    if fail_on_diff:
                        raise AssertionError(
                            f"Difference found at source_row_index {source_row_index}, target_row_index {target_row_index - 1 if not target_exhausted else 'N/A'}: {differences}"
                        )
                    result.append({"source_row_index": source_row_index, "differences": differences})

                source_row_index += 1

            # Report remaining target rows that do not have a matching source row.
            while True:
                try:
                    target_row = target.reader.get_row(row_idx=target_row_index, resource=target.resource)
                except StopIteration:
                    break

                target_dict = target_row.get_dict_row_data(target_header_map)
                differences: ColumnDifference = {
                    h: {"source": None, "target": target_dict.get(h)}
                    for h in compare_headers
                    if target_dict.get(h) is not None
                }

                if differences:
                    if fail_on_diff:
                        raise AssertionError(
                            f"Difference found at source_row_index N/A, target_row_index {target_row_index}: {differences}"
                        )
                    result.append({"source_row_index": source_row_index, "differences": differences})

                source_row_index += 1
                target_row_index += 1

            return result
        finally:
            if self is not target:
                target.close()
