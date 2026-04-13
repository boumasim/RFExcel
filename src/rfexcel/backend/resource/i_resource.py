from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any

from rfexcel.model.cell_data.i_raw_cell_data import IRawCellData
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.utils.types import ColumnValues, InsertNativeType


class IResource(ABC):
	def __init__(self, path: Path):
		self._path: Path = path

	@property
	@abstractmethod
	def active_sheets(self) -> Any:
		"""Returns the active sheets in the resource, which may be a single sheet or a collection of sheets depending on the implementation."""
		pass

	@property
	@abstractmethod
	def current_sheet(self) -> str:
		"""Returns the name of the currently active sheet in the resource."""
		pass

	@property
	@abstractmethod
	def last_read_row_index(self) -> int:
		"""Returns the index of the last row that was read from the resource, which can be used for tracking progress when reading data."""
		pass

	@property
	def path(self) -> Path:
		"""Returns the file path associated with the resource, which can be used for reference or when saving changes to the file."""
		return self._path

	@abstractmethod
	def close(self):
		"""Closes the resource and releases any associated resources, such as file handles or memory."""
		pass

	@abstractmethod
	def get_sheet_names(self) -> list[str]:
		"""Retrieves a list of sheet names from the resource, allowing clients to identify and access specific sheets within the tabular file."""
		pass

	@abstractmethod
	def switch_sheet(self, name: str) -> None:
		"""Switches the active sheet in the resource to the specified sheet name, allowing clients to read from or write to different sheets within the tabular file."""
		pass

	@abstractmethod
	def fetch_row(self, row_index: int, **kwargs: Any) -> IRawRowData:
		"""Fetches a single row of data from the resource based on the specified row index, returning it as an IRawRowData object for further processing."""
		pass

	@abstractmethod
	def fetch_cell(self, cell_name: str, **kwargs: Any) -> IRawCellData:
		"""Fetches a single cell of data from the resource based on the specified cell name (e.g., "A1"), returning it as an IRawCellData object for further processing."""
		pass

	@abstractmethod
	def add_sheet(self, name: str) -> None:
		"""Adds a new sheet with the specified name to the resource, allowing clients to create new sheets within the tabular file for organizing data."""
		pass

	@abstractmethod
	def delete_sheet(self, name: str) -> None:
		"""Deletes the sheet with the specified name from the resource, allowing clients to remove sheets that are no longer needed within the tabular file."""
		pass

	@abstractmethod
	def save(self, path: Path | None = None) -> None:
		"""Saves the current state of the resource to the specified file path, or overwrites the existing file if no path is provided."""
		pass

	@abstractmethod
	def append_row(self, cell_data: ColumnValues) -> None:
		"""Appends a new row of data to the end of the current sheet in the resource, using the provided cell data for the new row."""
		pass

	@abstractmethod
	def update_row(self, row_index: int, cell_data: ColumnValues) -> None:
		"""Updates an existing row of data in the resource at the specified row index, using the provided cell data to overwrite the existing values in that row."""
		pass

	@abstractmethod
	def delete_row(self, row_index: int) -> None:
		"""Deletes a row of data from the resource at the specified row index, allowing clients to remove rows that are no longer needed within the current sheet."""
		pass

	@abstractmethod
	def insert_row(self, row_index: int, cell_data: ColumnValues) -> None:
		"""Inserts a new row of data into the resource at the specified row index, using the provided cell data for the new row and shifting existing rows down as needed."""
		pass

	@abstractmethod
	def set_cell(self, cell_name: str, value: InsertNativeType) -> None:
		"""Sets the value of a specific cell in the resource identified by its name (e.g., "A1") to the provided value, allowing clients to update individual cells within the current sheet."""
		pass
