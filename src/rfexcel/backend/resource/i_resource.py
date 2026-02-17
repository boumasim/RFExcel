from abc import ABC, abstractmethod, abstractproperty

from rfexcel.utlis.types import Row


class IResource(ABC):

    @property
    @abstractmethod
    def header_row(self) -> int:
        """Return the 1-based row number where headers are located."""
        pass

    @abstractmethod
    def close(self):
        pass

    @abstractmethod
    def get_row(self, row_index: int) -> Row:
        """Return a single row by index (0-based). 
        
        For streaming resources, must be called sequentially. Attempting to read
        a previously read row will raise an exception.
        
        Args:
            row_index: The row index (0-based, data rows only, headers already read)
            
        Returns:
            Dictionary representing the row with headers as keys
            
        Raises:
            IndexError: If row_index is out of bounds
            RuntimeError: If trying to read backwards in streaming mode
        """
        pass
