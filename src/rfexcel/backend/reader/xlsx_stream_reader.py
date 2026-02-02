from typing import override

from openpyxl import Workbook
from .i_reader import IReader


class XlsxStreamReader(IReader):

    def __init__(self, wb: Workbook):
        self._wb = wb

    @override
    def print(self):
        print("xlsx stream reader\n")