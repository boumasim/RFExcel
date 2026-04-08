from __future__ import annotations

import weakref
from pathlib import Path
from typing import override

from rfexcel.advice.interceptors import auto_convert_xls_to_xlsx
from rfexcel.backend.interfaces.i_library import ISetExcel
from rfexcel.backend.resource.i_resource import IResource
from rfexcel.backend.writer.i_writer import IWriter
from rfexcel.exception.library_exceptions import OperationNotSupportedForFormat
from rfexcel.utils.types import ColumnValues, InsertNativeType


class XlsWriter(IWriter):

    def set_excel_reference(self, ref: ISetExcel):
        self._ref = weakref.ref(ref)

    def resolve_weak_ref(self) -> ISetExcel:
        ref = self._ref()
        if ref is None:
            raise ReferenceError("RFExcel reference is not set or has been garbage collected.")
        return ref

    @override
    def print(self):
        print("xls writer\n")

    @override
    @auto_convert_xls_to_xlsx
    def add_sheet(self, name: str, resource: IResource):
        raise OperationNotSupportedForFormat("Write operations not supported for xls format")

    @override
    @auto_convert_xls_to_xlsx
    def delete_sheet(self, name: str, resource: IResource):
        raise OperationNotSupportedForFormat("Write operations not supported for xls format")

    @override
    @auto_convert_xls_to_xlsx
    def save(self, path: Path | None, resource: IResource) -> None:
        raise OperationNotSupportedForFormat("Write operations not supported for xls format")

    @override
    @auto_convert_xls_to_xlsx
    def append_row(self, cell_data: ColumnValues, resource: IResource) -> None:
        raise OperationNotSupportedForFormat("Write operations not supported for xls format")

    @override
    @auto_convert_xls_to_xlsx
    def update_row(self, row_index: int, cell_data: ColumnValues, resource: IResource) -> None:
        raise OperationNotSupportedForFormat("Write operations not supported for xls format")

    @override
    @auto_convert_xls_to_xlsx
    def delete_row(self, row_index: int, resource: IResource) -> None:
        raise OperationNotSupportedForFormat("Write operations not supported for xls format")

    @override
    @auto_convert_xls_to_xlsx
    def insert_row(self, row_index: int, cell_data: ColumnValues, resource: IResource) -> None:
        raise OperationNotSupportedForFormat("Write operations not supported for xls format")

    @override
    @auto_convert_xls_to_xlsx
    def set_cell(self, cell_name: str, value: InsertNativeType, resource: IResource) -> None:
        raise OperationNotSupportedForFormat("Write operations not supported for xls format")