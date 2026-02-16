from .backend.metadata.i_metadata import IMetadata
from .backend.metadata.null_metadata import NullMetadata
from .backend.reader.i_reader import IReader
from .backend.reader.null_reader import NullReader
from .backend.resource.i_resource import IResource
from .backend.resource.null_resource import NullResource
from .backend.style.i_style import IStyle
from .backend.style.null_style import NullStyle
from .backend.writer.i_writer import IWriter
from .backend.writer.null_writer import NullWriter
from .utlis.types import Data


class RFExcel:

    def __init__(self,
                writer: IWriter = NullWriter(),
                reader: IReader = NullReader(),
                style: IStyle = NullStyle(),
                metadata: IMetadata = NullMetadata(),
                resource: IResource = NullResource()):
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

    def get_rows(self) -> Data:
        """Get all rows from the workbook as a list of dictionaries.
        
        The first row in the workbook is treated as column headers.
        Each subsequent row is returned as a dictionary where keys are column headers.
        
        Returns:
            List[Dict[str, str]]: List of rows, each row is a dictionary.
        """
        return self._reader.get_rows(self._resource)