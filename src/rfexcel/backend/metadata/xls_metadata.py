from typing_extensions import override

from rfexcel.backend.metadata.i_metadata import IMetadata
from rfexcel.backend.resource.i_resource import IResource


class XlsMetadata(IMetadata):

    def __init__(self):
        pass

    @override
    def print(self):
        print("xls metadata\n")

    @override
    def get_sheet_names(self, resource: IResource) -> list[str]:
        return resource.get_sheet_names()