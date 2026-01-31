from typing import override
from .i_style import IStyle


class XlsxStyle(IStyle):

    @override
    def print(self) -> None:
        print("xlsx style\n")