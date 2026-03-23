from pathlib import Path
from typing import override
import weakref

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.backend.writer.i_writer import IWriter
from rfexcel.RFExcel import RFExcel
from rfexcel.utlis.types import ColumnValues


class XlsWriter(IWriter):

    def set_excel_reference(self, ref: RFExcel):
        self._ref = weakref.ref(ref)

    def _resolve_weak_ref(self) -> RFExcel:
        ref = self._ref()
        if ref is None:
            raise ReferenceError("RFExcel reference is not set or has been garbage collected.")
        return ref

    def _convert_to_xlsx(self):
        self._resolve_weak_ref().xls_to_xlsx()

    @override
    def print(self):
        print("xls writer\n")

    @override
    def add_sheet(self, name: str, resource: IResource):
        self._convert_to_xlsx()
        ref = self._resolve_weak_ref()
        ref.add_sheet(name=name)

    @override
    def delete_sheet(self, name: str, resource: IResource):
        self._convert_to_xlsx()
        ref = self._resolve_weak_ref()
        ref.delete_sheet(name)

    @override
    def save(self, path: Path | None, resource: IResource) -> None:
        self._convert_to_xlsx()
        ref = self._resolve_weak_ref()
        ref.save_workbook(str(path) if path else None)

    @override
    def append_row(self, cell_data: ColumnValues, resource: IResource) -> None:
        self._convert_to_xlsx()
        ref = self._resolve_weak_ref()
        ref.writer.append_row(cell_data, ref.resource)

    @override
    def update_row(self, row_index: int, cell_data: ColumnValues, resource: IResource) -> None:
        self._convert_to_xlsx()
        ref = self._resolve_weak_ref()
        ref.writer.update_row(row_index, cell_data, ref.resource)

    @override
    def delete_row(self, row_index: int, resource: IResource) -> None:
        self._convert_to_xlsx()
        ref = self._resolve_weak_ref()
        ref.writer.delete_row(row_index, ref.resource)

    @override
    def insert_row(self, row_index: int, cell_data: ColumnValues, resource: IResource) -> None:
        self._convert_to_xlsx()
        ref = self._resolve_weak_ref()
        ref.writer.insert_row(row_index, cell_data, ref.resource)