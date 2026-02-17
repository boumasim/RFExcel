from typing_extensions import override
from xlrd import Book

from rfexcel.backend.reader.i_reader import IReader
from rfexcel.backend.resource.i_resource import IResource
from rfexcel.utlis.types import Data


class XlsOnDemandReader(IReader):

    def __init__(self):
        pass

    @override
    def print(self):
        print("xls on_demand reader\n")