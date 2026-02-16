from typing import override

from rfexcel.exception.library_exceptions import LibraryException
from rfexcel.utlis.types import Row

from .i_resource import IResource


class NullResource(IResource):

    @override
    def close(self):
        raise LibraryException("Invalid operation: resource not available")

    @override
    def get_row(self, row_index: int) -> Row:
        raise LibraryException("Invalid operation: resource not available")