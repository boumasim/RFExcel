from typing import override

from openpyxl import Workbook

from .i_reader import IReader


class XlsxStreamReader(IReader):

    def __init__(self):
        pass

    @override
    def print(self):
        print("xlsx stream reader\n")