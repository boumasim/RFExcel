from typing import override

from openpyxl import Workbook
from .i_resource import IResource


class XlsxResource(IResource):

    def __init__(self, wb: Workbook):
        self._wb: Workbook = wb

    @override
    def close(self):
        self._wb.close()