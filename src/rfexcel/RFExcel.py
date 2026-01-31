from .backend.style.i_style import IStyle
from .backend.reader.i_reader import IReader
from .backend.writer.i_writer import IWriter
from .backend.metadata.i_metadata import IMetadata

class RFExcel:
    _writer: IWriter
    _reader: IReader
    _style: IStyle
    _metadata: IMetadata

    def __init__(self, writer: IWriter, reader: IReader, style: IStyle, metadata: IMetadata):
        self._writer = writer
        self._reader = reader
        self._style = style
        self._metadata = metadata

    def print(self):
        self._writer.print()
        self._reader.print()
        self._style.print()
        self._metadata.print()