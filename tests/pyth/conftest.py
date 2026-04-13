from typing import Generator

import pytest

from rfexcel.RFExcelLibrary import RFExcelLibrary

@pytest.fixture
def lib() -> Generator[RFExcelLibrary, None, None]:
    library = RFExcelLibrary()
    yield library
    library.close()

