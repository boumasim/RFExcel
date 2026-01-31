from typing import override
from .i_writer import IWriter


class NullWriter(IWriter):

    @override
    def print(self):
        print("will throw writer exception\n")