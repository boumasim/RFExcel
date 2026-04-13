from collections.abc import Generator

import pytest

from rfexcel.RFExcelLibrary import RFExcelLibrary


@pytest.fixture
def lib() -> Generator[RFExcelLibrary]:
    library = RFExcelLibrary()
    yield library
    library.close()