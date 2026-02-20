from typing_extensions import override

from rfexcel.backend.metadata.i_metadata import IMetadata


class XlsMetadata(IMetadata):

    def __init__(self):
        pass

    @override
    def print(self):
        print("xls metadata\n")