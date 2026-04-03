from datetime import datetime, timedelta
from typing import TypeAlias, TypedDict

NativeType: TypeAlias = str | int | float | bool | datetime | timedelta | None # Underlying types used by libraries
InsertNativeType: TypeAlias = str | int | float | bool
CellValue: TypeAlias = NativeType
InsertDictType: TypeAlias = dict[str, InsertNativeType]

ListRowData: TypeAlias = list[NativeType]             # A row as a plain list of cell values
HeaderMap: TypeAlias = dict[str, int]                 # {header_name : column_index}
DictRowData: TypeAlias = dict[str, NativeType]        # {header_name : cell value}
HeaderSpec: TypeAlias = HeaderMap | list[str]
ColumnValues: TypeAlias = dict[int, InsertNativeType]       # {column_index : value} used for inserting

# Support types for compare_data_to
class ValueDifference(TypedDict):
    source: NativeType
    target: NativeType
ColumnDifference: TypeAlias = dict[str, ValueDifference]
class RowDifference(TypedDict):
    source_row_index: int
    target_row_index: int
    differences: ColumnDifference
