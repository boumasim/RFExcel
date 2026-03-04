from typing import Any, List, Union

from robot.api import logger  # type: ignore
from robot.api.deco import keyword, not_keyword  # type: ignore
from robot.utils import DotDict  # type: ignore

from rfexcel.backend.lib.i_library import IExcel
from rfexcel.factory.workbook_factory import WorkbookFactory
from rfexcel.utlis.types import DictRowData, ListRowData


class RFExcelLibrary:

    ROBOT_LIBRARY_SCOPE = "TEST CASE"
    ROBOT_LIBRARY_LISTENER = "SELF"
    ROBOT_LISTENER_API_VERSION = 2

    def __init__(self):
        self._factory = WorkbookFactory()
        self._active_workbook: IExcel | None = None

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
    def get_rows(self,
                header_row: int = 1,
                search_criteria: dict[str, str] | str | None = None, 
                partial_match: bool = False,
                one_row: bool = False,
                **kwargs: Any) -> List[DictRowData] | DictRowData:
        """Returns all data rows from the active workbook as a list of dictionaries.

        The row specified by ``header_row`` is used as the column header. Every
        subsequent row is returned as a ``dict`` where keys are the header values
        and values are the corresponding cell contents. All values are returned
        as strings.

        Rows *before* ``header_row`` are ignored entirely. If ``header_row``
        points beyond the last row of the file, an empty list is returned.

        *Filtering with search_criteria*

        When ``search_criteria`` is provided, only rows where *all* rules match
        are returned (AND logic). Each rule is a key-value pair where the key is
        a column header and the value is what to match against.

        ``search_criteria`` can be supplied in two forms:
        - A ``dict``: ``{"Product ID": "P-200", "Price": "25.50"}``
        - A string with ``key=value`` pairs separated by ``;``:
            ``"Product ID=P-200;Price=25.50"``

        When ``partial_match=True``, each criteria value only needs to be a
        *substring* of the corresponding cell value (e.g. ``"Keyboard"`` matches
        ``"Keyboard, Mechanical"``).
        When ``partial_match=False`` *(default)*, each cell value must equal
        the criteria value exactly.

        If any key in ``search_criteria`` is not present in the column headers,
        no row can satisfy that criterion and the result will be empty.

        *Returning a single row*

        When ``one_row=True``, iteration stops at the first matching row and that
        row is returned directly as a flat ``dict`` rather than a list. If no row
        matches, an empty ``dict`` is returned.

        *Backend keyword arguments*

        Any additional ``**kwargs`` are forwarded all the way to the underlying
        library calls (openpyxl / xlrd / csv). For example:
        - ``data_only=True`` *(xlsx)* — return cached cell values instead of formulas.
        - ``data_only=False`` *(xlsx)* — return formula strings as-is.

        In streaming / on-demand mode (``.xlsx`` read-only, ``.csv`` read-only),
        rows are consumed sequentially. Calling ``Get Rows`` a second time on the
        same open workbook will raise a ``StreamingViolationException``.
        In edit mode, repeated calls return the full data set each time.

        Returns an empty list (or empty dict when ``one_row=True``) if:
        - No workbook is currently open.
        - The file is empty or contains only a header row.
        - ``header_row`` is beyond the last row of the file.
        - ``search_criteria`` is provided but no row matches.

        Arguments:
        - ``header_row``: Row number (1-based) to use as column headers. Defaults to ``1``.
        - ``search_criteria``: Optional filter. Dict or ``"key=value;key=value"`` string.
        - ``partial_match``: If ``True``, criteria values are matched as substrings. Defaults to ``False``.
        - ``one_row``: If ``True``, return the first matching row as a flat dict. Defaults to ``False``.
        - ``**kwargs``: Forwarded to the backend library (e.g. ``data_only=True`` for xlsx).

        Examples:
        | Load Workbook | ${CURDIR}/data.xlsx |                                         |                      |
        | ${rows} =     | Get Rows            |                                         |                      |
        | ${rows} =     | Get Rows            | search_criteria=Product ID=P-200        |                      |
        | ${rows} =     | Get Rows            | search_criteria=Description=Keyboard    | partial_match=True   |
        | ${row} =      | Get Rows            | search_criteria=Product ID=P-200        | one_row=True         |
        | ${rows} =     | Get Rows            | search_criteria=${dict}                 |                      |
        | ${rows} =     | Get Rows            | header_row=2                            |                      |
        """
        if self._active_workbook:
            return self._active_workbook.get_rows(
                header_row=header_row,
                search_criteria=search_criteria,
                partial_match=partial_match,
                one_row=one_row,
                **kwargs,
            )
        return DotDict() if one_row else []

    @keyword("Get Row")  # pyright: ignore[reportUntypedFunctionDecorator]
    def get_row(self, row: int, headers: ListRowData | None = None, **kwargs: Any) -> Union[DictRowData, ListRowData]:
        """Returns a single row from the active workbook.

        The ``row`` argument is 1-based. The ``headers`` argument controls the
        return format:

        - *No headers (default)*: Returns the row as a plain ``list`` of string
            values. Useful for positional access.
        - *With headers*: Maps the row values to the provided header names and
            returns a ``dict``, identical in structure to a row returned by
            ``Get Rows``.

        Any additional ``**kwargs`` are forwarded to the underlying backend library
        (e.g. ``data_only=True`` for xlsx to get cached cell values).

        Returns an empty list if no workbook is open or the row index is beyond
        the last row of the file.

        Arguments:
        - ``row``: Row number to read (1-based).
        - ``headers``: Optional list of column names to map values against.
        - ``**kwargs``: Forwarded to the backend library.

        Examples:
        | Load Workbook  | ${CURDIR}/data.xlsx |                               |               |
        | ${row} =       | Get Row             | 2                             |               |
        | Log            | ${row}[0]           |                               |               |
        | ${headers} =   | Create List         | Name | Age | Country           |               |
        | ${row} =       | Get Row             | 2    | headers=${headers}        |               |
        | Log            | ${row}[Name]        |                               |               |
        """
        resolved: list[str] = headers if headers is not None else []
        if self._active_workbook: return self._active_workbook.get_row(row=row, headers=resolved, **kwargs)
        return []

    @keyword("List Sheet Names")  # pyright: ignore[reportUntypedFunctionDecorator]
    def list_sheet_names(self) -> list[str]:
        """Returns the names of all sheets in the active workbook.

        Works for ``.xlsx`` and ``.xls`` formats (both edit and streaming modes).
        Raises ``OperationNotSupportedForFormat`` when called on a CSV workbook,
        as CSV files do not have the concept of sheets.

        Returns an empty list if no workbook is currently open.

        Examples:
        | Load Workbook       | ${CURDIR}/data.xlsx  |
        | ${sheets} =         | List Sheet Names     |
        | Should Contain      | ${sheets}            | Sheet1 |
        | Load Workbook       | ${CURDIR}/data.xls   |
        | ${sheets} =         | List Sheet Names     |
        """
        if self._active_workbook:
            return self._active_workbook.list_sheet_names()
        return []

    @keyword("Switch Sheet")  # pyright: ignore[reportUntypedFunctionDecorator]
    def switch_sheet(self, name: str) -> None:
        """Switches the active sheet within the currently open workbook.

        Supported for ``.xlsx`` and ``.xls`` formats in all modes.
        Raises ``OperationNotSupportedForFormat`` when called on a CSV workbook.
        Raises ``LibraryException`` if no workbook is currently open.

        Arguments:
        - ``name``: The exact name of the sheet to activate.

        Examples:
        | Load Workbook  | ${CURDIR}/data.xlsx |        |
        | Switch Sheet   | Sheet2              |        |
        | ${rows} =      | Get Rows            |        |
        | Load Workbook  | ${CURDIR}/data.xls  |        |
        | Switch Sheet   | Second              |        |
        """
        if self._active_workbook:
            self._active_workbook.switch_sheet(name)

    @keyword("Add Sheet")  # pyright: ignore[reportUntypedFunctionDecorator]
    def add_sheet(self, name: str) -> None:
        """Adds a new sheet with the given name to the active workbook and switches to it.

        The new sheet becomes the active sheet immediately after creation, so
        subsequent read/write operations will target the newly added sheet.

        Supported formats and modes:
        - ``.xlsx`` (edit mode): Full support.
        - ``.xls`` (edit mode): The file is *lazily converted* to ``.xlsx`` format
          in memory before the sheet is added. The original ``.xls`` file on disk
          is *not* modified.
        - ``.xlsx`` (streaming mode): Raises ``NotSupportedInReadOnlyMode``.
        - ``.xls`` (streaming/on-demand mode): Not supported; raises ``OperationNotSupportedForFormat``.
        - ``.csv``: Raises ``OperationNotSupportedForFormat`` — CSV files have no concept of sheets.

        Raises ``LibraryException`` if no workbook is currently open.

        Arguments:
        - ``name``: The name to assign to the new sheet.

        Examples:
        | Load Workbook | ${CURDIR}/data.xlsx |          |
        | Add Sheet     | NewSheet            |          |
        | ${sheets} =   | List Sheet Names    |          |
        | Should Contain | ${sheets}          | NewSheet |
        """
        if self._active_workbook:
            self._active_workbook.add_sheet(name)

    @keyword("Delete Sheet")  # pyright: ignore[reportUntypedFunctionDecorator]
    def delete_sheet(self, name: str) -> None:
        """Deletes the sheet with the given name from the active workbook.

        After deletion, the active sheet is reset to the first remaining sheet
        in the workbook.

        Supported formats and modes:
        - ``.xlsx`` (edit mode): Full support.
        - ``.xls`` (edit mode): The file is *lazily converted* to ``.xlsx`` format
          in memory before the sheet is deleted. The original ``.xls`` file on disk
          is *not* modified.
        - ``.xlsx`` (streaming mode): Raises ``LibraryException``.
        - ``.xls`` (streaming/on-demand mode): Raises ``LibraryException``.
        - ``.csv``: Raises ``OperationNotSupportedForFormat`` — CSV files have no concept of sheets.

        Raises ``LibraryException`` if the sheet does not exist or no workbook is open.

        Arguments:
        - ``name``: The exact name of the sheet to delete.

        Examples:
        | Load Workbook | ${CURDIR}/data.xlsx |          |
        | Delete Sheet  | OldSheet            |          |
        | ${sheets} =   | List Sheet Names    |          |
        | Should Not Contain | ${sheets}      | OldSheet |
        """
        if self._active_workbook:
            self._active_workbook.delete_sheet(name)

    @keyword("Switch Source")  # pyright: ignore[reportUntypedFunctionDecorator]
    def switch_source(self, path: str, read_only: bool = False, **kwargs: Any) -> None:
        """Switches the active workbook to a different file.

        This is a convenience method that combines ``Close Workbook`` and
        ``Load Workbook`` into a single step. It first closes the currently
        active workbook (if any), then opens the new file specified by
        ``source``.

        Arguments:
        - ``path``: Path to the new workbook to load.
        - ``read_only``: Whether to open the workbook in read-only mode. Defaults to ``False``.
        - ``**kwargs``: Additional keyword arguments passed to ``Load Workbook``.

        Examples:
        | Switch Source | ${CURDIR}/data.xlsx | read_only=True |
        | Switch Source | ${CURDIR}/data.csv  | read_only=True |
        """
        self.close()
        self.load_workbook(path=path, read_only=read_only, **kwargs)