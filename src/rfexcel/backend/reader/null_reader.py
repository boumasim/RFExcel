from typing import override

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.exception.library_exceptions import LibraryException
from rfexcel.utlis.types import Data

from .i_reader import IReader


class NullReader(IReader):

    @override
    def print(self):
        raise LibraryException("Invalid operation: reader not available")