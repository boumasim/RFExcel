from typing import override

from rfexcel.backend.writer.i_writer import IWriter


class CsvWriter(IWriter):
    def __init__(self):
        pass

    @override
    def print(self):
        print("csv reader")