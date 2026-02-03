from typing_extensions import override
from xlrd import Book

from rfexcel.backend.reader.i_reader import IReader


class XlsOnDemandReader(IReader):

    def __init__(self, wb: Book):
        self._wb: Book = wb

    @override
    def print(self):
        print("xls on_demand reader\n")