from typing import override

from rfexcel.backend.metadata.i_metadata import IMetadata
from rfexcel.backend.resource.i_resource import IResource
from rfexcel.exception.library_exceptions import OperationNotSupportedForFormat


class CsvMetadata(IMetadata):

    @override
    def print(self) -> None:
        print("No metadata available for CSV format")

    @override
    def get_sheet_names(self, resource: IResource) -> list[str]:
        raise OperationNotSupportedForFormat("CSV format does not support multiple sheets")