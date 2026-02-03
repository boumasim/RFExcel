from pathlib import Path

import xlrd
from openpyxl import Workbook
from openpyxl.reader import excel
from robot.api import logger

from rfexcel.RFExcel import RFExcel
from rfexcel.backend.metadata.xls_metadata import XlsMetadata
from rfexcel.backend.metadata.xlsx_metadata import XlsxMetadata
from rfexcel.backend.reader.xls_on_demand_reader import XlsOnDemandReader
from rfexcel.backend.reader.xls_standart_reader import XlsStandardReader
from rfexcel.backend.reader.xlsx_edit_reader import XlsxEditReader
from rfexcel.backend.reader.xlsx_stream_reader import XlsxStreamReader
from rfexcel.backend.resource.xls_resource import XlsResource
from rfexcel.backend.resource.xlsx_resource import XlsxResource
from rfexcel.backend.style.xls_style import XlsStyle
from rfexcel.backend.style.xlsx_style import XlsxStyle
from rfexcel.backend.writer.null_writer import NullWriter
from rfexcel.backend.writer.xlsx_writer import XlsxWriter
from rfexcel.exception.library_exceptions import FileAlreadyExistsException, FileFormatNotSupportedException
from rfexcel.exception.library_exceptions import FileDoesNotExistException
from rfexcel.rfexcel_constants import CSV_SUFFIX, VALID_SUFFIXES, XLS_SUFFIX, XLSX_SUFFIX


class WorkbookFactory:

    def create_workbook(self, path: str, **kwargs) -> RFExcel:
        file_path: Path = Path(path)
        extension: str = file_path.suffix.lower()

        if extension not in VALID_SUFFIXES: raise FileFormatNotSupportedException()
        if file_path.exists(): raise FileAlreadyExistsException()

        file_path.parent.mkdir(parents=True, exist_ok=True)

        if extension == XLSX_SUFFIX:
            return self._create_xlsx_edit(path = file_path, **kwargs)
        if extension == XLS_SUFFIX:
            raise FileFormatNotSupportedException(msg="Use xlsx format for creating and editing excel files.")
        else:
            raise Exception("Exception in create_workbook occured")


    def load_workbook(self, path: str, read_only: bool = False, **kwargs):
        file_path: Path = Path(path)
        extension: str = file_path.suffix.lower()

        if extension not in VALID_SUFFIXES: raise FileFormatNotSupportedException()
        if not file_path.exists(): raise FileDoesNotExistException(file_path.name)

        if extension == XLSX_SUFFIX and read_only:
            return self._load_xlsx_stream(path=file_path, **kwargs)
        elif extension == XLSX_SUFFIX and not read_only:
            return self._load_xlsx_edit(path=file_path, **kwargs)
        elif extension == XLS_SUFFIX and read_only:
            return self._load_xls_on_demand(path=file_path, **kwargs)
        elif extension == XLS_SUFFIX and not read_only:
            return self._load_xls_standard(path=file_path, **kwargs)
        elif extension == CSV_SUFFIX and read_only:
            pass
        elif extension == CSV_SUFFIX and not read_only:
            pass
        else:
            raise Exception("Exception in load_workbook occured")

    def _create_xlsx_edit(self, path: Path, **kwargs) -> RFExcel:
        wb: Workbook = Workbook(**kwargs)
        ws = wb.active
        assert ws is not None
        ws.title = path.name.lower()
        wb.save(filename=path)
        return RFExcel(XlsxWriter(wb=wb), XlsxEditReader(wb=wb), XlsxStyle(wb=wb), XlsxMetadata(wb=wb), XlsxResource(wb=wb))

    def _load_xlsx_stream(self, path: Path, **kwargs) -> RFExcel:
        _data_only: bool = kwargs.get('data_only', False)
        wb: Workbook = excel.load_workbook(filename=path, read_only=True, data_only=_data_only)
        return RFExcel(NullWriter(), XlsxStreamReader(wb), XlsxStyle(wb), XlsxMetadata(wb), XlsxResource(wb))

    def _load_xlsx_edit(self, path: Path, **kwargs) -> RFExcel:
        wb: Workbook = excel.load_workbook(filename=path, read_only=False, **kwargs)
        return RFExcel(XlsxWriter(wb), XlsxEditReader(wb), XlsxStyle(wb), XlsxMetadata(wb), XlsxResource(wb))

    def _load_xls_on_demand(self, path: Path, **kwargs) -> RFExcel:
        _formating_info: bool = kwargs.get('formatting_info', True)
        wb: xlrd.Book = xlrd.open_workbook(str(path), on_demand=True, formatting_info=_formating_info, **kwargs)
        return RFExcel(NullWriter(), XlsOnDemandReader(wb), XlsStyle(wb), XlsMetadata(wb), XlsResource(wb))

    def _load_xls_standard(self, path: Path, **kwargs) -> RFExcel:
        _formating_info: bool = kwargs.get('formatting_info', True)
        wb: xlrd.Book = xlrd.open_workbook(str(path), on_demand=False, formatting_info=_formating_info, **kwargs)
        return RFExcel(NullWriter(), XlsStandardReader(wb), XlsStyle(wb), XlsMetadata(wb), XlsResource(wb))