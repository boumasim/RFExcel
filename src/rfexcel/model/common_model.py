

from typing import Any

import xlrd
from xlrd.sheet import Cell

def norm_xls_value(cell: Cell) -> Any:
    """Converts value to library friendly type"""
    if cell.ctype == xlrd.XL_CELL_NUMBER:
        v: float = float(cell.value)
        if v.is_integer():
            return int(v)
    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
        return bool(cell.value)
    elif cell.ctype == xlrd.XL_CELL_EMPTY or cell.ctype == xlrd.XL_CELL_BLANK or cell.ctype == xlrd.XL_CELL_ERROR:
        return ""
    return cell.value