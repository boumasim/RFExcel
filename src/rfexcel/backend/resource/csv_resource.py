import csv
from io import TextIOWrapper
from itertools import zip_longest
from pathlib import Path
from typing import Dict, List, override

from robot.api import logger

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.exception.library_exceptions import (RowIndexOutOfBoundsException,
                                                  StreamingViolationException)
from rfexcel.rfexcel_constants import BASE_DIALECT, BASE_ENCODING
from rfexcel.utlis.types import Data, Row


class CsvEditResource(IResource):
    def __init__(self, path: Path, header_row: int = 1, dialect: str = BASE_DIALECT, encoding: str = BASE_ENCODING, **kwargs):
        self._path = path
        self._encoding = encoding
        self._edited = False
        self._header_row = header_row
        
        with open(path, mode='r', newline='', encoding=encoding) as f:
            all_rows = list(csv.reader(f, dialect=dialect, **kwargs))
            
            header_index = header_row - 1
            if len(all_rows) > header_index:
                self._fieldnames = all_rows[header_index]
                data_rows = all_rows[header_index + 1:]
                self._data: Data = [dict(zip(self._fieldnames, row)) for row in data_rows]
            else:
                self._fieldnames = []
                self._data = []

    @property
    @override
    def header_row(self) -> int:
        """Return the 1-based row number where headers are located."""
        return self._header_row
                
    @override
    def get_row(self, row_index: int) -> Row:
        """Returns row at given index (1-based data index).
        
        Arguments:
        - ``row_index``: 1-based index of data row (row 1 is first data row after headers)
        """
        list_index = row_index - 1
        
        if list_index < 0 or list_index >= len(self._data):
            raise StopIteration()
            
        return self._data[list_index]

    @override
    def close(self):
        if self._edited:
            with open(self._path, mode='w', newline='', encoding=self._encoding) as f:
                if not self._fieldnames and self._data:
                    self._fieldnames = list(self._data[0].keys())

                writer = csv.DictWriter(f, fieldnames=self._fieldnames)
                writer.writeheader()
                writer.writerows(self._data)
            logger.info(f"CSV file '{self._path.name}' was updated.")


class CsvStreamResource(IResource):
    def __init__(self, path: Path, header_row: int = 1, dialect: str = BASE_DIALECT, encoding: str = BASE_ENCODING, **kwargs):
        self._path = path
        self._handle = open(path, mode='r', newline='', encoding=encoding)
        self._reader = csv.reader(self._handle, dialect=dialect, **kwargs)
        self._header_row = header_row
        
        self._headers: list[str] = []
        for i in range(header_row):
            try:
                row = next(self._reader)
                if i == header_row - 1:
                    self._headers = row
            except StopIteration:
                break
        
        self._last_read_data_index = 0

    @property
    @override
    def header_row(self) -> int:
        """Return the 1-based row number where headers are located."""
        return self._header_row

    @override
    def get_row(self, row_index: int) -> Row:
        """Returns row at given index (1-based data index).
        
        Arguments:
        - ``row_index``: 1-based index of data row.
        """
        
        if row_index <= self._last_read_data_index:
            raise StreamingViolationException(row_index, self._last_read_data_index)
        
        try:
            while self._last_read_data_index < row_index - 1:
                next(self._reader)
                self._last_read_data_index += 1
            
            raw_row = next(self._reader)
            self._last_read_data_index = row_index
            
            row_dict: Row = {}
            for header, value in zip_longest(self._headers, raw_row, fillvalue=""):
                if header:
                    row_dict[header] = str(value) if value else ""
            return row_dict

        except StopIteration:
            raise StopIteration()

    @override
    def close(self):
        if self._handle and not self._handle.closed:
            self._handle.close()