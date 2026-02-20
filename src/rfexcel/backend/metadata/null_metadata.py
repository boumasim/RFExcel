from typing import override
from .i_metadata import IMetadata


class NullMetadata(IMetadata):

    @override
    def print(self):
        print("will throw metadata exception\n")