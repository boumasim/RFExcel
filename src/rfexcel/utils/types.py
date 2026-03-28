from typing import TypeAlias, TypedDict

ListRowData: TypeAlias = list[str]             # A row as a plain list of string values
HeaderMap: TypeAlias = dict[str, int]          # {header_name : column_index}
DictRowData: TypeAlias = dict[str, str]        # User-supplied {header_name : value} for writes / search
HeaderSpec: TypeAlias = HeaderMap | list[str]
ColumnValues: TypeAlias = dict[int, str]       # Internal {column_index : value} passed to writers/resources


# Support types for compare_data_to
class ValueDifference(TypedDict):
    source: str
    target: str
ColumnDifference: TypeAlias = dict[str, ValueDifference]
class RowDifference(TypedDict):
    source_row_index: int
    differences: ColumnDifference