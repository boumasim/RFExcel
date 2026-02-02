from typing import override

from openpyxl import Workbook
from .i_writer import IWriter


class XlsxWriter(IWriter):

    def __init__(self, wb: Workbook):
        self._wb = wb

    @override
    def print(self):
        print("xlsx writer\n")