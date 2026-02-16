from typing_extensions import override
from xlrd import Book

from rfexcel.backend.reader.i_reader import IReader


class XlsStandardReader(IReader):

    def __init__(self):
        pass

    @override
    def print(self):
        print("xls standard reader\n")