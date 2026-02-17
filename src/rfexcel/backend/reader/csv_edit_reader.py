from typing import override

from rfexcel.backend.reader.i_reader import IReader
from rfexcel.backend.resource.i_resource import IResource
from rfexcel.utlis.types import Data


class CsvEditReader(IReader):
    def __init__(self):
        pass

    @override
    def print(self):
        print("csv edit reader")