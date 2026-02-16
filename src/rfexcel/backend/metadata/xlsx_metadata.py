from typing import override

from openpyxl import Workbook

from .i_metadata import IMetadata


class XlsxMetadata(IMetadata):

    def __init__(self):
        pass

    @override
    def print(self):
        print("xlsx metadata\n")