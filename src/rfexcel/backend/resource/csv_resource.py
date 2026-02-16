import csv
from io import TextIOWrapper
from pathlib import Path
from typing import override

from robot.api import logger

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.rfexcel_constants import BASE_DIALECT, BASE_ENCODING
from rfexcel.utlis.types import CsvData


class CsvEditResource(IResource):
    def __init__(self, path: Path, dialect: str = BASE_DIALECT, encoding: str = BASE_ENCODING, **kwargs):
        self._path = path
        self._edited = False
        with open(path, mode='r', newline='', encoding=encoding) as f:
            reader = csv.DictReader(f, dialect=dialect)
            self._fieldnames = reader.fieldnames
            self._data: CsvData = list(reader)

    @override
    def close(self):
        if self._edited:
            with open(self._path, mode='w', newline='') as f:
                fieldnames = self._fieldnames if self._fieldnames is not None else []
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(self._data)
            logger.info("Csv file was edited, file rewriten")

class CsvStreamResource(IResource):
    def __init__(self, path: Path, dialect: str = BASE_DIALECT, encoding: str = BASE_ENCODING, **kwargs):
        self._path = path
        self._handle = open(path, mode='r', newline='', encoding=encoding)
        self._reader = csv.reader(self._handle, dialect=dialect, **kwargs)

    @override
    def close(self):
        if self._handle and not self._handle.closed:
            self._handle.close()