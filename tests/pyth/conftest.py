from pathlib import Path
from typing import Generator

import pytest

from rfexcel.RFExcelLibrary import RFExcelLibrary

_RESOURCES = Path(__file__).parent.parent / "resources"

XLSX_FILE = str(_RESOURCES / "data.xlsx")
CSV_FILE  = str(_RESOURCES / "data.csv")
XLS_FILE  = str(_RESOURCES / "example.xls")


@pytest.fixture
def lib() -> Generator[RFExcelLibrary, None, None]:
    library = RFExcelLibrary()
    yield library
    if library._active_workbook:
        library.close()

