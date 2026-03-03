from typing import override

from rfexcel.backend.resource.i_resource import IResource

from .i_metadata import IMetadata


class XlsxMetadata(IMetadata):

    def __init__(self):
        pass

    @override
    def print(self):
        print("xlsx metadata\n")

    @override
    def get_sheet_names(self, resource: IResource) -> list[str]:
        return resource.get_sheet_names()