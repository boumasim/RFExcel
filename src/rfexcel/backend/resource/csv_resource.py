import csv
from io import TextIOWrapper
from typing import override

from robot.api import logger

from rfexcel.backend.resource.i_resource import IResource
from rfexcel.utlis.types import CsvData


class CsvEditResource(IResource):
    def __init__(self, path: str):
        self._path = path
        self._edited = False
        with open(path, mode='r', newline='') as f:
            reader = csv.DictReader(f)
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
    def __init__(self, handle: TextIOWrapper):
        self._handle = handle

    @override
    def close(self):
        self._handle.close()