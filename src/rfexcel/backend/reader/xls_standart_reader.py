from typing_extensions import override
from xlrd import Book

from rfexcel.backend.reader.i_reader import IReader
from rfexcel.backend.resource.i_resource import IResource
from rfexcel.utlis.types import Data


class XlsStandardReader(IReader):

    def __init__(self):
        pass

    @override
    def print(self):
        print("xls standard reader\n")

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