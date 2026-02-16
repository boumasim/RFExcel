from typing import override

from openpyxl import Workbook

from .i_reader import IReader


class XlsxEditReader(IReader):

    def __init__(self):
        pass

    @override
    def print(self):
        print("xlsx edit reader\n")