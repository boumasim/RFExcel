from robot.api import logger  # type: ignore
from robot.api.deco import keyword, not_keyword  # type: ignore

from rfexcel.RFExcel import RFExcel
from rfexcel.factory.workbook_factory import WorkbookFactory


class RFExcelLibrary:

    ROBOT_LIBRARY_SCOPE = "TEST CASE"
    ROBOT_LIBRARY_LISTENER = "SELF"
    ROBOT_LISTENER_API_VERSION = 2

    def __init__(self):
        self._factory = WorkbookFactory()
        self._active_workbook: RFExcel | None = None

    @not_keyword
    def end_test(self, name, attrs):
        logger.info("Cleanup after test execution...")
        if self._active_workbook: self.close()

    @keyword("Create Workbook")
    def create_workbook(self, path: str, **kwargs) -> None:
        """
        Creates a new workbook at the given path.

        *Arguments:*
        - ``path``: Path where the new file will be created (including extension, e.g., .xlsx).
        - ``kwargs``: Additional optional parameters for the backend.

        *Example:*
        | Create Workbook | ${OUTPUT_DIR}${/}result.xlsx |
        """
        self._active_workbook = self._factory.create_workbook(path=path, **kwargs)
        logger.info("Workbook successfully created")

    @keyword("Load Workbook")
    def load_workbook(self, path: str, read_only: bool = False, **kwargs) -> None:
        """
        Opens an existing workbook.

        If ``read_only=True``, the file is opened in **streaming mode**.
        This mode is memory efficient for reading large files but does not support writing.

        *Arguments:*
        - ``path``: Path to the existing file.
        - ``read_only``: If set to True, opens the file in read-only (stream) mode. Defaults to False.

        *Examples:*

        | Load Workbook | data.xlsx |   |
        | Load Workbook | large_dataset.xlsx | read_only=True |
        """
        self._active_workbook = self._factory.load_workbook(path=path, read_only=read_only, **kwargs)
        logger.info("Workbook successfully opened")

    @keyword("Print")
    def print(self):
        if self._active_workbook: self._active_workbook.print()

    @keyword("Close Workbook")
    def close(self):
        """
        Closes the active workbook.

        This keyword is called **automatically** at the end of a test case (via Listener)
        if the workbook was not closed manually.
        """
        if self._active_workbook: self._active_workbook.close()
        logger.info("File successfully closed")
        self._active_workbook = None
