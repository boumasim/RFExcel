from typing import override
from .i_reader import IReader


class XlsxStreamReader(IReader):

    @override
    def print(self):
        print("xlsx stream reader\n")