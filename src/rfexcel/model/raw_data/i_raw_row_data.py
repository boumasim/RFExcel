from abc import ABC, abstractmethod

from rfexcel.utils.types import DictRowData, HeaderMap, ListRowData


class IRawRowData(ABC):
    """
    Unified contract for extracting and normalizing row data from underlying engines
    (csv, xlrd for .xls, openpyxl for .xlsx).
    
    This interface ensures that regardless of the source file format, the Robot Framework
    facade receives structurally identical data, perfectly aligned native Python types,
    and consistent handling of empty or missing cells.
    """

    @abstractmethod
    def get_list_row_data(self) -> ListRowData:
        """
        Returns the physical row as a list of natively typed values.
        
        - Numeric strings and .0 floats MUST be safely cast to `int` or `float`.
        - Ignores empty or None values.
        """
        pass

    @abstractmethod
    def get_dict_row_data(self, header_map: HeaderMap) -> DictRowData:
        """
        Returns normalized row data as a dictionary mapped to the provided headers.
        
        - Uses 1-based column indices provided by the `header_map`.
        - Existing, but empty cells MUST return `""`.
        - Requested columns that are out-of-bounds MUST return `""`.
        """
        pass

    @abstractmethod
    def get_header_map(self) -> HeaderMap:
        """
        Parses the row to create a column mapping for future lookups.
        
        - Returns a dictionary of {stripped_header_name: 1-based_column_index}.
        - Empty cells or whitespace-only strings MUST be ignored.
        """
        pass