from typing import override

from rfexcel.backend.writer.i_writer import IWriter
from rfexcel.utlis.types import CsvData


class CsvWriter(IWriter):
    def __init__(self, data: CsvData):
        pass

    @override
    def print(self):
        print("csv reader")