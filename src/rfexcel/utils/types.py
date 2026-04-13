from datetime import datetime, timedelta
from typing import TypedDict

type NativeType = (
	str | int | float | bool | datetime | timedelta | None
)  # Underlying types used by libraries
type InsertNativeType = str | int | float | bool
type CellValue = NativeType
type InsertDictType = dict[str, InsertNativeType]

type ListRowData = list[NativeType]  # A row as a plain list of cell values
type HeaderMap = dict[str, int]  # {header_name : column_index}
type DictRowData = dict[str, NativeType]  # {header_name : cell value}
type HeaderSpec = HeaderMap | list[str]
type ColumnValues = dict[int, InsertNativeType]  # {column_index : value} used for inserting


# Support types for compare_data_to
class ValueDifference(TypedDict):
    source: NativeType
    target: NativeType
type ColumnDifference = dict[str, ValueDifference]
class RowDifference(TypedDict):
    source_row_index: int
    target_row_index: int
    differences: ColumnDifference
