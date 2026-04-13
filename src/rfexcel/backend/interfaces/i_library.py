from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any

from rfexcel.backend.reader.i_reader import IReader
from rfexcel.backend.resource.i_resource import IResource
from rfexcel.backend.writer.i_writer import IWriter
from rfexcel.utils.types import (
	DictRowData,
	HeaderSpec,
	InsertDictType,
	InsertNativeType,
	ListRowData,
	NativeType,
	RowDifference,
)


class IExcel(ABC):
	@property
	@abstractmethod
	def read_only(self) -> bool:
		"""Indicates whether the Excel file is opened in read-only mode."""
		pass

	@property
	@abstractmethod
	def resource(self) -> IResource:
		"""Provides access to the underlying resource representing the tabular file."""
		pass

	@property
	@abstractmethod
	def reader(self) -> IReader:
		"""Provides access to the reader component responsible for reading data from the tabular file."""
		pass

	@property
	@abstractmethod
	def writer(self) -> IWriter:
		"""Provides access to the writer component responsible for writing data to the tabular file."""
		pass

	@abstractmethod
	def close(self):
		"""Closes the Excel file and releases any associated resources."""
		pass

	@abstractmethod
	def get_rows(
		self,
		header_row: int,
		search_criteria: str | dict[str, str] | None = None,
		partial_match: bool = False,
		one_row: bool = False,
		**kwargs: Any,
	) -> list[DictRowData] | DictRowData:
		"""Retrieves rows from the tabular file based on the specified criteria, using the provided header row for mapping."""
		pass

	@abstractmethod
	def list_sheet_names(self) -> list[str]:
		"""Returns a list of sheet names available in the Excel file."""
		pass

	@abstractmethod
	def switch_sheet(self, name: str) -> None:
		"""Switches the active sheet to the specified sheet name."""
		pass

	@abstractmethod
	def add_sheet(self, name: str) -> None:
		"""Adds a new sheet with the specified name to the tabular file."""
		pass

	@abstractmethod
	def delete_sheet(self, name: str) -> None:
		"""Deletes the sheet with the specified name from the tabular file."""
		pass

	@abstractmethod
	def get_row(self, row: int, headers: HeaderSpec, **kwargs: Any) -> DictRowData | ListRowData:
		"""Retrieves a single row of data from the tabular file based on the specified row number and header specification."""
		pass

	@abstractmethod
	def get_cell(self, cell_name: str) -> NativeType:
		"""Retrieves the value of a specific cell identified by its name (e.g., 'A1', 'B2') from the tabular file."""
		pass

	@abstractmethod
	def set_cell(self, cell_name: str, value: InsertNativeType) -> None:
		"""Sets the value of a specific cell identified by its name (e.g., 'A1', 'B2') in the tabular file."""
		pass

	@abstractmethod
	def save_workbook(self, path: str | None = None) -> None:
		"""Saves the current state of the workbook to the specified path. If no path is provided, it saves to the original location."""
		pass

	@abstractmethod
	def append_row(self, row_data: InsertDictType, header_row: int) -> None:
		"""Appends a new row of data to the end of the sheet, using the specified header row for mapping."""
		pass

	@abstractmethod
	def append_rows(self, rows: list[InsertDictType], header_row: int) -> None:
		"""Appends multiple rows of data to the end of the sheet, using the specified header row for mapping."""
		pass

	@abstractmethod
	def update_values(
		self,
		search_criteria: str | dict[str, str],
		values: InsertDictType,
		header_row: int,
		partial_match: bool,
		first_only: bool,
	) -> int:
		"""Updates values in rows that match the specified search criteria, using the provided header row for mapping."""
		pass

	@abstractmethod
	def delete_rows(
		self,
		search_criteria: str | dict[str, str],
		header_row: int,
		partial_match: bool,
		first_only: bool,
	) -> int:
		"""Deletes rows that match the specified search criteria, using the provided header row for mapping."""
		pass

	@abstractmethod
	def delete_row(self, row_number: int) -> None:
		"""Deletes a single row identified by its row number from the sheet."""
		pass

	@abstractmethod
	def insert_row(self, row_data: InsertDictType, row: int, header_row: int) -> None:
		"""Inserts a new row of data at the specified row number, using the provided header row for mapping."""
		pass

	@abstractmethod
	def compare_data_to(
		self,
		target: IExcel,
		source_header_row: int,
		target_header_row: int,
		target_sheet: str | None,
		headers: list[str] | None,
		fail_on_diff: bool,
	) -> list[RowDifference]:
		"""Compares the data in the current Excel file to another target tabular file, based on the specified header rows and optional sheet and header filters."""
		pass

class ISetExcel(ABC):
    @abstractmethod
    def xls_to_xlsx(self):
		"""Converts an XLS file to XLSX format, returning a new IExcel instance representing the converted file."""
		pass

	@property
	@abstractmethod
	def writer(self) -> IWriter:
		"""Provides access to the writer component responsible for writing data to the tabular file, with support for format conversion if necessary."""
        pass