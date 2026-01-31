from .backend.style.i_style import IStyle
from .backend.reader.i_reader import IReader
from .backend.writer.i_writer import IWriter
from .backend.metadata.i_metadata import IMetadata

class RFExcel:

    def __init__(self, writer: IWriter, reader: IReader, style: IStyle, metadata: IMetadata):
        self._writer: IWriter = writer
        self._reader: IReader = reader
        self._style: IStyle = style
        self._metadata: IMetadata = metadata

    def print(self):
        self._writer.print()
        self._reader.print()
        self._style.print()
        self._metadata.print()