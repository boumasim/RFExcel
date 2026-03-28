from datetime import datetime, timedelta
from typing import Any, TypeAlias, TypedDict

CellValue: TypeAlias = str | int | float | bool | datetime | timedelta | None  # Normalised cell value across xlrd and openpyxl

ListRowData: TypeAlias = list[Any]             # A row as a plain list of cell values
HeaderMap: TypeAlias = dict[str, int]          # {header_name : column_index}
DictRowData: TypeAlias = dict[str, Any]        # {header_name : cell value}
HeaderSpec: TypeAlias = HeaderMap | list[str]
ColumnValues: TypeAlias = dict[int, str]       # Internal {column_index : value} passed to writers/resources


# Support types for compare_data_to
class ValueDifference(TypedDict):
    source: Any
    target: Any
ColumnDifference: TypeAlias = dict[str, ValueDifference]
class RowDifference(TypedDict):
    source_row_index: int
    differences: ColumnDifference
