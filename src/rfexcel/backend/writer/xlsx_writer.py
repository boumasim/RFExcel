from typing import override
from .i_writer import IWriter


class XlsxWriter(IWriter):

    @override
    def print(self):
        print("xlsx writer\n")