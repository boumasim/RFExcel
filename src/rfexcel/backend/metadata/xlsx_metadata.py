from typing import override
from .i_metadata import IMetadata


class XlsxMetadata(IMetadata):

    @override
    def print(self):
        print("xlsx metadata\n")