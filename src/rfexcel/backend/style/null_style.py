from typing import override
from .i_style import IStyle


class NullStyle(IStyle):

    @override
    def print(self):
        print("will throw style exception\n")