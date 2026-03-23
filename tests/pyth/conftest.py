from pathlib import Path
from typing import Generator

import pytest

from rfexcel.RFExcelLibrary import RFExcelLibrary

_RESOURCES = Path(__file__).parent.parent / "resources"

XLSX_FILE  = str(_RESOURCES / "data.xlsx")
XLSX2_FILE = str(_RESOURCES / "data2.xlsx")
CSV_FILE   = str(_RESOURCES / "data.csv")
XLS_FILE   = str(_RESOURCES / "example.xls")


@pytest.fixture
def lib() -> Generator[RFExcelLibrary, None, None]:
    library = RFExcelLibrary()
    yield library
    library.close()  # safe to call when nothing is open


@pytest.fixture
def loaded_xlsx() -> Generator[RFExcelLibrary, None, None]:
    library = RFExcelLibrary()
    library.load_workbook(XLSX_FILE)
    yield library
    library.close()  # safe to call when nothing is open

