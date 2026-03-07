from pathlib import Path
from typing import override

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.backend.writer.i_writer import IWriter
from rfexcel.RFExcel import RFExcel


class XlsWriter(IWriter):

    def set_excel_reference(self, ref: RFExcel):
        self._ref = ref

    def _convert_to_xlsx(self):
        self._ref.xls_to_xlsx()

    @override
    def print(self):
        print("xls writer\n")

    @override
    def add_sheet(self, name: str, resource: IResource):
        self._convert_to_xlsx()
        self._ref.add_sheet(name=name)

    @override
    def delete_sheet(self, name: str, resource: IResource):
        self._convert_to_xlsx()
        self._ref.delete_sheet(name)

    @override
    def save(self, path: Path | None, resource: IResource) -> None:
        self._convert_to_xlsx()
        self._ref.save_workbook(str(path) if path else None)