from typing import override
from .i_reader import IReader


class NullReader(IReader):

    @override
    def print(self):
        print("will throw reader exception\n")