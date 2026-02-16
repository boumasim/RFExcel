from typing import override

from openpyxl import Workbook

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.utlis.types import Data

from .i_reader import IReader


class XlsxEditReader(IReader):

    def __init__(self):
        pass

    @override
    def print(self):
        print("xlsx edit reader\n")

    @override
    def get_rows(self, resource: IResource) -> Data:
        """Read all rows using get_row() in a loop."""
        result: Data = []
        row_index = 0
        
        while True:
            try:
                row = resource.get_row(row_index)
                result.append(row)
                row_index += 1
            except (IndexError, StopIteration):
                break
        
        return result