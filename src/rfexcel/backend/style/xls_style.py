from typing_extensions import override
from rfexcel.backend.style.i_style import IStyle


class XlsStyle(IStyle):

    def __init__(self):
        pass

    @override
    def print(self):
        print("xls style\n")