from typing import override

from rfexcel.backend.reader.i_reader import IReader


class CsvStreamReader(IReader):
    def __init__(self):
        pass

    @override
    def print(self):
        print("csv stream reader")