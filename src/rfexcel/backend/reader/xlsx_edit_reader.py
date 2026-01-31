from typing import override
from .i_reader import IReader


class XlsxEditReader(IReader):

    @override
    def print(self):
        print("xlsx edit reader\n")