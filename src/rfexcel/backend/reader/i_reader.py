from abc import ABC, abstractmethod
from typing import Any

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData


class IReader(ABC):
	@abstractmethod
	def get_headers(self, header_row_idx: int, resource: IResource, **kwargs: Any) -> IRawRowData:
		"""Retrieves the header row from the tabular file based on the specified header row index and resource."""
		pass

	@abstractmethod
	def get_row(self, row_idx: int, resource: IResource, **kwargs: Any) -> IRawRowData:
		"""Retrieves a single row from the tabular file based on the specified row index and resource."""
		pass