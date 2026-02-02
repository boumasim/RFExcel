from .backend.resource.i_resource import IResource
from .backend.style.i_style import IStyle
from .backend.reader.i_reader import IReader
from .backend.writer.i_writer import IWriter
from .backend.metadata.i_metadata import IMetadata

from openpyxl import Workbook

class RFExcel:

    def __init__(self, writer: IWriter, reader: IReader, style: IStyle, metadata: IMetadata, resource: IResource):
        self._writer: IWriter = writer
        self._reader: IReader = reader
        self._style: IStyle = style
        self._metadata: IMetadata = metadata
        self._resource: IResource = resource

    def print(self):
        self._writer.print()
        self._reader.print()
        self._style.print()
        self._metadata.print()

    def close(self):
        self._resource.close()