from abc import ABC, abstractmethod

from rfexcel.backend.resource.i_resource import IResource


class IMetadata(ABC):
	@abstractmethod
	def get_sheet_names(self, resource: IResource) -> list[str]:
		"""Retrieves a list of sheet names from the provided resource."""
		pass