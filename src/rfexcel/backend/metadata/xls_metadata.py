from typing_extensions import override
from xlrd import Book

from rfexcel.backend.metadata.i_metadata import IMetadata


class XlsMetadata(IMetadata):

    def __init__(self, wb: Book):
        self._wb: Book = wb

    @override
    def print(self):
        print("xls metadata\n")