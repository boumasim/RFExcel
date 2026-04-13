from abc import ABC, abstractmethod
from pathlib import Path

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.utils.types import ColumnValues, InsertNativeType


class IWriter(ABC):
	@abstractmethod
	def add_sheet(self, name: str, resource: IResource):
		"""Adds a new sheet with the specified name to the provided resource, allowing clients to create new sheets within the tabular file for organizing data."""
		pass

	@abstractmethod
	def delete_sheet(self, name: str, resource: IResource):
		"""Deletes the sheet with the specified name from the provided resource, allowing clients to remove sheets that are no longer needed within the tabular file."""
		pass

	@abstractmethod
	def save(self, path: Path | None, resource: IResource) -> None:
		"""Saves the current state of the provided resource to the specified file path, or overwrites the existing file if no path is provided."""
		pass

	@abstractmethod
	def append_row(self, cell_data: ColumnValues, resource: IResource) -> None:
		"""Appends a new row of data to the end of the current sheet in the provided resource, using the provided cell data for the new row."""
		pass

	@abstractmethod
	def update_row(self, row_index: int, cell_data: ColumnValues, resource: IResource) -> None:
		"""Updates an existing row of data in the provided resource at the specified row index, using the provided cell data to overwrite the existing values in that row."""
		pass

	@abstractmethod
	def delete_row(self, row_index: int, resource: IResource) -> None:
		"""Deletes a row of data from the provided resource at the specified row index, allowing clients to remove rows that are no longer needed within the current sheet."""
		pass

	@abstractmethod
	def insert_row(self, row_index: int, cell_data: ColumnValues, resource: IResource) -> None:
		"""Inserts a new row of data into the provided resource at the specified row index, using the provided cell data for the new row and shifting existing rows down as needed."""
		pass

	@abstractmethod
	def set_cell(self, cell_name: str, value: InsertNativeType, resource: IResource) -> None:
		"""Sets the value of a specific cell in the provided resource identified by its name (e.g., "A1") to the provided value, allowing clients to update individual cells within the current sheet."""
		pass
