from typing import override

from rfexcel.backend.resource.i_resource import IResource

from .i_writer import IWriter


class XlsxWriter(IWriter):

    def __init__(self):
        pass

    @override
    def print(self):
        print("xlsx writer\n")

    @override
    def add_sheet(self, name: str, resource: IResource):
        resource.add_sheet(name)