from typing import Any

from robot.api import logger  # type: ignore
from robot.api.deco import keyword, not_keyword  # type: ignore

from rfexcel.factory.workbook_factory import WorkbookFactory
from rfexcel.RFExcel import RFExcel
from rfexcel.utlis.types import Data


class RFExcelLibrary:

    ROBOT_LIBRARY_SCOPE = "TEST CASE"
    ROBOT_LIBRARY_LISTENER = "SELF"
    ROBOT_LISTENER_API_VERSION = 2

    def __init__(self):
        self._factory = WorkbookFactory()
        self._active_workbook: RFExcel | None = None

    @not_keyword  # pyright: ignore[reportUntypedFunctionDecorator]
    def end_test(self, name: str, attrs: dict[str, Any]) -> None:
        logger.info("Cleanup after test execution...")
        if self._active_workbook: self.close()

    @keyword("Create Workbook")  # pyright: ignore[reportUntypedFunctionDecorator]
    def create_workbook(self, path: str, **kwargs: Any) -> None:
        """Creates a new empty workbook at the given path and opens it in edit mode.

        Parent directories in the path are created automatically if they do not exist.
        The new file is immediately saved to disk and opened for editing.

        Supported formats:
        - ``.xlsx``: Creates a blank Excel workbook via openpyxl.
        - ``.csv``: Creates an empty CSV file.
        - ``.xls``: *Not supported* for creation. Use ``.xlsx`` instead.

        Raises ``FileAlreadyExistsException`` if a file at ``path`` already exists.
        Raises ``FileFormatNotSupportedException`` if the extension is unsupported or ``.xls``.

        Arguments:
        - ``path``: Destination path including the file extension (e.g., ``/tmp/result.xlsx``).

        Examples:
        | Create Workbook | ${OUTPUT_DIR}${/}result.xlsx |
        | Create Workbook | ${OUTPUT_DIR}${/}output.csv  |
        """
        self._active_workbook = self._factory.create_workbook(path=path, **kwargs)
        logger.info("Workbook successfully created")

    @keyword("Load Workbook")  # pyright: ignore[reportUntypedFunctionDecorator]
    def load_workbook(self, path: str, read_only: bool = False, **kwargs: Any) -> None:
        """Opens an existing workbook for reading or editing.

        The ``read_only`` flag controls which mode the file is opened in:

        - ``read_only=False`` *(default — Edit mode)*: Loads the entire file into memory.
          Supports both reading and writing. Suitable for small to medium-sized files.
        - ``read_only=True`` *(Streaming / On-Demand mode)*: Memory-efficient.
          Iterates over rows without loading the whole file. Supports reading only.
          For ``.xlsx`` and ``.csv`` this is strict forward-only access.
          For ``.xls`` this uses on-demand sheet loading (random row access is still available).

        Supported formats and modes:
        - ``.xlsx``: Edit and streaming mode.
        - ``.xls``: Read-only in both modes (``xlrd`` does not support writing).
        - ``.csv``: Edit and streaming mode.

        Optional keyword arguments (passed through to the backend):
        - ``data_only`` *(xlsx streaming only)*: If ``True``, formula cells return their
          last cached value instead of the formula string. Defaults to ``False``.

        Raises ``FileDoesNotExistException`` if the file cannot be found at ``path``.
        Raises ``FileFormatNotSupportedException`` for unsupported file extensions.

        Arguments:
        - ``path``: Path to the existing file.
        - ``read_only``: Open in streaming/on-demand mode if ``True``. Defaults to ``False``.

        Examples:
        | Load Workbook | ${CURDIR}/data.xlsx |                |                  |
        | Load Workbook | ${CURDIR}/large.xlsx | read_only=True |                  |
        | Load Workbook | ${CURDIR}/data.xls  |                |                  |
        | Load Workbook | ${CURDIR}/data.csv  | read_only=True |                  |
        | Load Workbook | ${CURDIR}/data.xlsx | read_only=True | data_only=True   |
        """
        self._active_workbook = self._factory.load_workbook(path=path, read_only=read_only, **kwargs)
        logger.info("Workbook successfully opened")

    @keyword("Print")  # pyright: ignore[reportUntypedFunctionDecorator]
    def print(self):
        if self._active_workbook: self._active_workbook.print()

    @keyword("Close Workbook")  # pyright: ignore[reportUntypedFunctionDecorator]
    def close(self) -> None:
        """Closes the active workbook and releases all associated resources.

        This releases any open file handles held by the backend (e.g., openpyxl
        read-only workbook connections, CSV file handles). After this keyword,
        a new ``Load Workbook`` or ``Create Workbook`` call is required before
        performing any further operations.

        This keyword is also called *automatically* at the end of every test case
        via the Robot Framework listener (``ROBOT_LIBRARY_LISTENER``), so explicit
        cleanup is not required but is recommended for clarity.

        Safe to call when no workbook is currently open — it will do nothing.

        Examples:
        | Load Workbook  | ${CURDIR}/data.xlsx |
        | ${rows} =      | Get Rows            |
        | Close Workbook |                     |
        """
        if self._active_workbook: self._active_workbook.close()
        logger.info("File successfully closed")
        self._active_workbook = None

    @keyword("Get Rows")  # pyright: ignore[reportUntypedFunctionDecorator]
    def get_rows(self, header_row: int = 1) -> Data:
        """Returns all data rows from the active workbook as a list of dictionaries.

        The row specified by ``header_row`` is used as the column header. Every
        subsequent row is returned as a ``dict`` where keys are the header values
        and values are the corresponding cell contents. All values are returned
        as strings.

        Rows *before* ``header_row`` are ignored entirely. If ``header_row``
        points beyond the last row of the file, an empty list is returned.

        In streaming / on-demand mode (``.xlsx`` read-only, ``.csv`` read-only),
        rows are consumed sequentially. Calling ``Get Rows`` a second time on the
        same open workbook will return an empty list because the stream is exhausted.
        In edit mode, repeated calls return the full data set each time.

        Returns an empty list if:
        - No workbook is currently open.
        - The file is empty or contains only a header row.
        - ``header_row`` is beyond the last row of the file.

        Arguments:
        - ``header_row``: Row number (1-based) to use as column headers. Defaults to ``1``.

        Examples:
        | Load Workbook | ${CURDIR}/data.xlsx |             |
        | ${rows} =     | Get Rows            |             |
        | Log           | ${rows[0]}          |             |
        | Load Workbook | ${CURDIR}/data.xlsx |             |
        | ${rows} =     | Get Rows            | header_row=2 |
        """
        if self._active_workbook: return self._active_workbook.get_rows(header_row=header_row)
        return []

