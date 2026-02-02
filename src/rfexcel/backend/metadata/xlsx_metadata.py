from typing import override

from openpyxl import Workbook
from .i_metadata import IMetadata


class XlsxMetadata(IMetadata):

    def __init__(self, wb: Workbook):
        self._wb = wb

    @override
    def print(self):
        print("xlsx metadata\n")