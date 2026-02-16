from typing import override

from openpyxl import Workbook

from .i_style import IStyle


class XlsxStyle(IStyle):

    def __init__(self):
        pass

    @override
    def print(self) -> None:
        print("xlsx style\n")