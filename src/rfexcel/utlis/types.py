from typing import List

from robot.utils import DotDict  # type: ignore

DictRowData = DotDict           # A row returned to the user: {column_header: cell_value}
ListRowData = List[str]         # A row as a plain list of string values
HeaderMap = dict[str, int]      # {header_name: 1-based_column_index}
RowInputData = dict[str, str]   # User-supplied {column_header: cell_value} for writes / search
HeaderSpec = HeaderMap | list[str]  # Accepted as headers param to Get Row
ColumnValues = dict[int, str]       # Internal {1-based_column_index: cell_value} passed to writers/resources
