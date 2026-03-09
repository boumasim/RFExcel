from typing import override

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.exception.library_exceptions import NullComponentException

from .i_metadata import IMetadata


class NullMetadata(IMetadata):

    @override
    def print(self):
        print("will throw metadata exception\n")

    @override
    def get_sheet_names(self, resource: IResource) -> list[str]:
        raise NullComponentException()