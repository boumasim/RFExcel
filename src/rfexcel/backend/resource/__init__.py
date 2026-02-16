"""Resource components for different file formats and modes.

Resources handle low-level file access and data extraction.
Each format has separate classes for edit and streaming modes.
"""

from .csv_resource import CsvEditResource, CsvStreamResource
from .i_resource import IResource
from .null_resource import NullResource
from .xls_resource import XlsEditResource, XlsStreamResource
from .xlsx_resource import XlsxEditResource, XlsxStreamResource

__all__ = [
    'IResource',
    'NullResource',
    'XlsxEditResource',
    'XlsxStreamResource',
    'XlsEditResource',
    'XlsStreamResource',
    'CsvEditResource',
    'CsvStreamResource',
]
