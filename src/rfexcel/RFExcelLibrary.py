from rfexcel.RFExcel import RFExcel
from rfexcel.factory.workbook_factory import WorkbookFactory

from robot.api.deco import keyword # type: ignore


class RFExcelLibrary:

    ROBOT_LIBRARY_SCOPE = "TEST CASE"

    def __init__(self):
        self._factory = WorkbookFactory()
        self._active_workbook: RFExcel | None = None

    @keyword("Create Workbook")
    def create_workbook(self, path: str, read_only: bool = False):
        self._active_workbook = self._factory.create_workbook(path=path, read_only=read_only)

    @keyword("Print")
    def print(self):
        if self._active_workbook: self._active_workbook.print()