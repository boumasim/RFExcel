from rfexcel.backend.metadata.null_metadata import NullMetadata
from rfexcel.backend.metadata.xlsx_metadata import XlsxMetadata
from rfexcel.backend.reader.null_reader import NullReader
from rfexcel.backend.reader.xlsx_edit_reader import XlsxEditReader
from rfexcel.backend.reader.xlsx_stream_reader import XlsxStreamReader
from rfexcel.backend.style.null_style import NullStyle
from rfexcel.backend.style.xlsx_style import XlsxStyle
from rfexcel.backend.writer.null_writer import NullWriter
from rfexcel.backend.writer.xlsx_writer import XlsxWriter
from rfexcel.RFExcel import RFExcel

class WorkbookFactory:

    def create_workbook(self, path: str, read_only: bool = False) -> RFExcel:
        if path.endswith(".xlsx") and not read_only: return self._create_xlsx_edit()
        elif path.endswith(".xlsx") and read_only: return self._create_xlsx_stream()
        else: return self._create_invalid_workbook()

    def load_workbook(self, path: str, read_only: bool = False):
        pass

    def _create_xlsx_stream(self) -> RFExcel:
        return RFExcel(NullWriter(), XlsxStreamReader(), NullStyle(), XlsxMetadata())

    def _create_xlsx_edit(self) -> RFExcel:
        return RFExcel(XlsxWriter(), XlsxEditReader(), XlsxStyle(), XlsxMetadata())

    def _create_invalid_workbook(self) -> RFExcel:
        return RFExcel(NullWriter(), NullReader(), NullStyle(), NullMetadata())