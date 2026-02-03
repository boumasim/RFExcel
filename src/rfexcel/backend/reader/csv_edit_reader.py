from typing import override

from rfexcel.backend.reader.i_reader import IReader
from rfexcel.utlis.types import CsvData


class CsvEditReader(IReader):
    def __init__(self, wb: CsvData):
        pass

    @override
    def print(self):
        print("csv edit reader")