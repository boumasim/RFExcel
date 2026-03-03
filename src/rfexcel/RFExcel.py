from typing import Any, List, Union

from robot.utils import DotDict  # type: ignore

from rfexcel.utlis.utilities import (convert_string_to_dict_row_data,
                                     search_in_row)

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
from .utlis.types import DictRowData, ListRowData


class RFExcel:

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

    def print(self):
        self._writer.print()
        self._reader.print()
        self._style.print()
        self._metadata.print()

    def close(self):
        self._resource.close()

    def get_rows(self, 
                header_row: int,
                search_criteria: str | dict[str, str] | None = None,
                partial_match: bool = False,
                one_row: bool = False,
                **kwargs: Any) -> List[DictRowData] | DictRowData:
        search_criteria_dict = convert_string_to_dict_row_data(search_criteria) if search_criteria else None

        try:
            headers = self._reader.get_headers(header_row_idx=header_row, resource=self._resource, **kwargs).get_list_row_data()
        except StopIteration:
            headers = []

        result: List[DictRowData] = []
        row_index = header_row + 1

        while True:
            try:
                row = self._reader.get_row(row_idx=row_index, resource=self._resource, **kwargs)
                if not search_criteria_dict or search_in_row(source_row=row.get_dict_row_data(headers=headers), search_criteria=search_criteria_dict, partial_match=partial_match):
                    result.append(row.get_dict_row_data(headers=headers))
                    if one_row:
                        break
                row_index += 1
            except StopIteration:
                break

        return result if not one_row else (result[0] if result else DotDict())

    def list_sheet_names(self) -> list[str]:
        return self._metadata.get_sheet_names(self._resource)

    def switch_sheet(self, name: str) -> None:
        self._resource.switch_sheet(name)

    def get_row(self, row: int, headers: list[str], **kwargs: Any) -> Union[DictRowData, ListRowData]:
        try:
            raw = self._reader.get_row(row_idx=row, resource=self._resource, **kwargs)
        except StopIteration:
            return []

        if not headers:
            return raw.get_list_row_data()
        return raw.get_dict_row_data(headers=headers)