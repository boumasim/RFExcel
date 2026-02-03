from typing import override

from xlrd import Book

from .i_resource import IResource


class XlsResource(IResource):

    def __init__(self, wb: Book):
        self._wb: Book = wb

    @override
    def close(self):
        self._wb.release_resources()