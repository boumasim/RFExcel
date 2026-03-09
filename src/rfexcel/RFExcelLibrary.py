from typing import Any, List, Union

from robot.api import logger  # type: ignore
from robot.api.deco import keyword, not_keyword  # type: ignore

from rfexcel.backend.lib.i_library import IExcel
from rfexcel.factory.workbook_factory import WorkbookFactory
from rfexcel.utlis.types import (DictRowData, HeaderSpec, ListRowData,
                                 RowInputData)


class RFExcelLibrary:
    """Robot Framework library for reading and writing Excel and CSV files.

    = Supported Formats =

    | Format    | Edit mode  | Streaming mode | Notes |
    | ``.xlsx`` | yes        | yes            | Full read/write via openpyxl. |
    | ``.xls``  | yes*       | yes*           | *Write operations trigger lazy in-memory conversion to ``.xlsx``; the original file is never modified. |
    | ``.csv``  | yes        | yes            | No sheet concept; sheet keywords raise ``OperationNotSupportedForFormat``. |

    = Modes =

    - *Edit mode* (``read_only=False``, default): Loads the full file into memory.
      Supports reading and writing.
    - *Streaming mode* (``read_only=True``): Memory-efficient, read-only.
      For ``.xlsx`` and ``.csv`` access is strictly forward-only — calling a read
      keyword twice on the same open workbook raises ``StreamingViolationException``.
      For ``.xls``, on-demand sheet loading is used; random row access is still available.

    = Search Criteria & Partial Matching =

    Keywords that filter or target rows accept a ``search_criteria`` argument.
    It can be supplied as:
    - A ``dict``: ``{"Product ID": "P-200", "Price": "25.50"}``
    - A string of ``key=value`` pairs separated by ``;``: ``"Product ID=P-200;Price=25.50"``

    Matching uses *AND* logic — all pairs must match for a row to qualify.
    A key absent from the column headers produces no match.

    When ``partial_match=True``, the criterion value only needs to be a *substring*
    of the cell value (e.g. ``"Keyboard"`` matches ``"Keyboard, Mechanical"``).
    Defaults to exact match (``partial_match=False``).
    """

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
        """Creates a new empty workbook at ``path`` and opens it in edit mode.

        Parent directories are created automatically. Supported: ``.xlsx``, ``.csv``.
        Raises ``FileAlreadyExistsException`` if a file already exists at ``path``.
        Raises ``FileFormatNotSupportedException`` for unsupported or ``.xls`` paths.

        Arguments:
        - ``path``: Destination path including the file extension.

        Examples:
        | Create Workbook | ${OUTPUT_DIR}${/}result.xlsx |
        | Create Workbook | ${OUTPUT_DIR}${/}output.csv  |
        """
        self._active_workbook = self._factory.create_workbook(path=path, **kwargs)
        logger.info("Workbook successfully created")

    @keyword("Load Workbook")  # pyright: ignore[reportUntypedFunctionDecorator]
    def load_workbook(self, path: str, read_only: bool = False, **kwargs: Any) -> None:
        """Opens an existing workbook for reading or editing.

        - ``read_only=False`` *(default — Edit mode)*: Loads the full file into memory; read/write.
        - ``read_only=True`` *(Streaming mode)*: Memory-efficient, read-only.
          For ``.xlsx`` and ``.csv`` access is strictly forward-only.
          For ``.xls`` on-demand sheet loading allows random row access.

        Raises ``FileDoesNotExistException`` if the file is not found at ``path``.
        Raises ``FileFormatNotSupportedException`` for unsupported file extensions.

        Arguments:
        - ``path``: Path to the existing file.
        - ``read_only``: Open in streaming mode if ``True``. Defaults to ``False``.
        - ``**kwargs``: Forwarded to the backend (e.g. ``data_only=True`` for xlsx streaming).

        Examples:
        | Load Workbook | ${CURDIR}/data.xlsx |                |                |
        | Load Workbook | ${CURDIR}/large.xlsx | read_only=True |                |
        | Load Workbook | ${CURDIR}/data.xls  |                |                |
        | Load Workbook | ${CURDIR}/data.csv  | read_only=True |                |
        | Load Workbook | ${CURDIR}/data.xlsx | read_only=True | data_only=True |
        """
        self._active_workbook = self._factory.load_workbook(path=path, read_only=read_only, **kwargs)
        logger.info("Workbook successfully opened")

    @keyword("Close Workbook")  # pyright: ignore[reportUntypedFunctionDecorator]
    def close(self) -> None:
        """Closes the active workbook and releases all associated resources.

        Called automatically at the end of each test case; explicit cleanup is
        optional but recommended for clarity. Safe to call with no open workbook.

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
                search_criteria: RowInputData | str | None = None,
                partial_match: bool = False,
                one_row: bool = False,
                **kwargs: Any) -> List[DictRowData] | DictRowData:
        """Returns data rows from the active sheet as a list of dicts keyed by column header.

        Row ``header_row`` is used as column headers; rows before it are ignored.
        See the `library introduction`_ for details on ``search_criteria`` and ``partial_match``.

        When ``one_row=True``, stops at the first matching row and returns it as a flat
        dict instead of a list. Returns ``{}`` if nothing matches.

        In streaming mode, rows are consumed sequentially — calling ``Get Rows`` twice
        on the same open workbook raises ``StreamingViolationException``.

        Returns ``[]`` (or ``{}`` when ``one_row=True``) if no workbook is open,
        ``header_row`` is beyond the file, or no row matches.

        Arguments:
        - ``header_row``: Row that contains the column headers (row 1 = first row). Defaults to ``1``.
        - ``search_criteria``: Optional filter — see `library description`_ for format details.
        - ``partial_match``: Substring matching when ``True`` — see `library description`_. Defaults to ``False``.
        - ``one_row``: Return first match as a flat dict when ``True``. Defaults to ``False``.
        - ``**kwargs``: Forwarded to the backend (e.g. ``data_only=True`` for xlsx).

        Examples:
        | Load Workbook | ${CURDIR}/data.xlsx |                                         |                    |
        | ${rows} =     | Get Rows            |                                         |                    |
        | ${rows} =     | Get Rows            | search_criteria=Product ID=P-200        |                    |
        | ${rows} =     | Get Rows            | search_criteria=Description=Keyboard    | partial_match=True |
        | ${row} =      | Get Rows            | search_criteria=Product ID=P-200        | one_row=True       |
        | ${rows} =     | Get Rows            | search_criteria=${dict}                 |                    |
        | ${rows} =     | Get Rows            | header_row=2                            |                    |
        """
        if self._active_workbook:
            return self._active_workbook.get_rows(
                header_row=header_row,
                search_criteria=search_criteria,
                partial_match=partial_match,
                one_row=one_row,
                **kwargs,
            )
        return DictRowData() if one_row else []

    @keyword("Get Row")  # pyright: ignore[reportUntypedFunctionDecorator]
    def get_row(self, row: int, headers: HeaderSpec | None = None, **kwargs: Any) -> Union[DictRowData, ListRowData]:
        """Returns a single row by its row number as a list or dict.

        - No ``headers``: Returns a plain ``list`` of string values.
        - ``headers`` as a list: Maps values by position to the given column names; returns a ``dict``.
        - ``headers`` as a dict ``{"Name": 2, "Age": 3}``: Uses column indices for lookup; returns a ``dict``.
          Use this for tables that do not start at column A.

        Returns ``[]`` if no workbook is open or the row is beyond the last row.
        Any ``**kwargs`` are forwarded to the backend (e.g. ``data_only=True`` for xlsx).

        Arguments:
        - ``row``: Row number to read (row 1 = first row).
        - ``headers``: Optional list of column names or dict mapping column names to column numbers.
        - ``**kwargs``: Forwarded to the backend.

        Examples:
        | Load Workbook  | ${CURDIR}/data.xlsx |                               |               |
        | ${row} =       | Get Row             | 2                             |               |
        | Log            | ${row}[0]           |                               |               |
        | ${headers} =   | Create List         | Name | Age | Country           |               |
        | ${row} =       | Get Row             | 2    | headers=${headers}        |               |
        | Log            | ${row}[Name]        |                               |               |
        | ${hmap} =      | Create Dictionary   | Name=2 | Age=3 | Country=4     |               |
        | ${row} =       | Get Row             | 2    | headers=${hmap}           |               |
        """
        resolved: HeaderSpec = headers if headers is not None else []
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
        """Adds a new sheet to the active workbook and switches to it.

        Supported in ``.xlsx`` (edit) and ``.xls`` (edit; lazily converted to ``.xlsx`` in memory).
        Streaming mode and ``.csv`` raise ``NotSupportedInReadOnlyMode`` or ``OperationNotSupportedForFormat``.

        Arguments:
        - ``name``: Name of the new sheet.

        Examples:
        | Load Workbook  | ${CURDIR}/data.xlsx |          |
        | Add Sheet      | NewSheet            |          |
        | ${sheets} =    | List Sheet Names    |          |
        | Should Contain | ${sheets}           | NewSheet |
        """
        if self._active_workbook:
            self._active_workbook.add_sheet(name)

    @keyword("Delete Sheet")  # pyright: ignore[reportUntypedFunctionDecorator]
    def delete_sheet(self, name: str) -> None:
        """Deletes a sheet from the active workbook; the first remaining sheet becomes active.

        Supported in ``.xlsx`` (edit) and ``.xls`` (edit; lazily converted to ``.xlsx`` in memory).
        Streaming mode and ``.csv`` raise ``LibraryException`` or ``OperationNotSupportedForFormat``.
        Raises ``LibraryException`` if the sheet does not exist.

        Arguments:
        - ``name``: The exact name of the sheet to delete.

        Examples:
        | Load Workbook      | ${CURDIR}/data.xlsx |          |
        | Delete Sheet       | OldSheet            |          |
        | ${sheets} =        | List Sheet Names    |          |
        | Should Not Contain | ${sheets}           | OldSheet |
        """
        if self._active_workbook:
            self._active_workbook.delete_sheet(name)

    @keyword("Save Workbook")  # pyright: ignore[reportUntypedFunctionDecorator]
    def save_workbook(self, path: str | None = None) -> None:
        """Saves the current state of the active workbook to disk.

        By default saves back to the original path. Providing ``path`` enables a
        *Save As* workflow and updates the active path for subsequent saves.

        Streaming / read-only mode raises ``NotSupportedInReadOnlyMode``.
        For ``.xls`` without a prior write operation, raises ``OperationNotSupportedForFormat``;
        trigger any write (e.g. ``Add Sheet``) first, then save to a ``.xlsx`` path.
        Safe to call when no workbook is open — does nothing.

        Arguments:
        - ``path``: Optional destination path. Omit to save to the original path.

        Examples:
        | Load Workbook  | ${CURDIR}/data.xlsx          |                              |
        | Add Sheet      | Report                       |                              |
        | Save Workbook  |                              |                              |
        | Save Workbook  | ${OUTPUT_DIR}${/}result.xlsx |                              |
        | Load Workbook  | ${CURDIR}/data.xls           |                              |
        | Add Sheet      | NewSheet                     |                              |
        | Save Workbook  | ${OUTPUT_DIR}${/}result.xlsx |                              |
        """
        if self._active_workbook:
            self._active_workbook.save_workbook(path=path)
            logger.info("Workbook successfully saved")

    @keyword("Append Row")  # pyright: ignore[reportUntypedFunctionDecorator]
    def append_row(self, row_data: RowInputData, header_row: int = 1) -> None:
        """Appends a new row to the end of the active sheet.

        ``row_data`` maps column header names to values. Keys not found in the headers
        are silently ignored; missing columns are written as empty strings.
        Streaming / read-only mode raises ``LibraryException``.
        ``.xls`` edit mode triggers lazy conversion to ``.xlsx`` in memory.

        Arguments:
        - ``row_data``: Dict mapping column header names to cell values.
        - ``header_row``: Row that contains the column headers (row 1 = first row). Defaults to ``1``.

        Examples:
        | Load Workbook | ${CURDIR}/data.xlsx |                                     |              |
        | Append Row    | ${{{"Product ID": "P-999", "Description": "Widget", "Price": "9.99", "Location": "Online"}}} |
        | Save Workbook |                     |                                     |              |
        | Load Workbook | ${CURDIR}/data.csv  |                                     |              |
        | Append Row    | ${{{"Product ID": "P-100", "Price": "1.00"}}}           |              |
        | Save Workbook |                     |                                     |              |
        """
        if self._active_workbook:
            self._active_workbook.append_row(row_data=row_data, header_row=header_row)

    @keyword("Append Rows")  # pyright: ignore[reportUntypedFunctionDecorator]
    def append_rows(self, rows: list[RowInputData], header_row: int = 1) -> None:
        """Appends multiple rows to the end of the active sheet. Same rules as ``Append Row``.

        Arguments:
        - ``rows``: List of dicts, each mapping column header names to cell values.
        - ``header_row``: Row that contains the column headers (row 1 = first row). Defaults to ``1``.

        Examples:
        | Load Workbook | ${CURDIR}/data.xlsx |                                                                         |
        | ${row1} =     | Create Dictionary   | Product ID=P-001 | Description=Gadget | Price=4.99 |
        | ${row2} =     | Create Dictionary   | Product ID=P-002 | Description=Widget | Price=9.99 |
        | Append Rows   | ${[${row1}, ${row2}]} |                                                               |
        | Save Workbook |                     |                                                                         |
        """
        if self._active_workbook:
            self._active_workbook.append_rows(rows=rows, header_row=header_row)

    @keyword("Update Values")  # pyright: ignore[reportUntypedFunctionDecorator]
    def update_values(self,
                      search_criteria: RowInputData | str,
                      values: RowInputData | str,
                      header_row: int = 1,
                      partial_match: bool = False,
                      first_only: bool = False) -> int:
        """Updates cells in all rows matching ``search_criteria``. Returns the count of updated rows.

        Only columns listed in ``values`` are overwritten; others are left untouched.
        Keys in ``values`` not present in headers are silently ignored.
        Streaming / read-only mode raises ``LibraryException``.
        See the `library introduction`_ for details on ``search_criteria`` and ``partial_match``.

        Arguments:
        - ``search_criteria``: Filter identifying which rows to update — see `library description`_ for format details.
        - ``values``: Dict of ``{column_header: new_value}`` pairs to write.
        - ``header_row``: Row that contains the column headers (row 1 = first row). Defaults to ``1``.
        - ``partial_match``: Substring matching when ``True`` — see `library description`_. Defaults to ``False``.
        - ``first_only``: Update only the first matching row when ``True``. Defaults to ``False``.

        Examples:
        | Load Workbook  | ${CURDIR}/data.xlsx |                                                    |                    |
        | ${count} =     | Update Values       | ${{{'Product ID': 'P-001'}}} | ${{{'Price': '0.00', 'Location': 'Archived'}}} |
        | Should Be Equal As Integers | ${count} | 1 |                                             |
        | Save Workbook  |                     |                                                    |                    |
        | Load Workbook  | ${CURDIR}/data.csv  |                                                    |                    |
        | Update Values  | Location=Online     | ${{{'Price': '0.00'}}}       | partial_match=True | first_only=True |
        | Save Workbook  |                     |                                                    |                    |
        """
        if self._active_workbook:
            return self._active_workbook.update_values(
                search_criteria=search_criteria,
                values=values,
                header_row=header_row,
                partial_match=partial_match,
                first_only=first_only,
            )
        return 0

    @keyword("Delete Rows")  # pyright: ignore[reportUntypedFunctionDecorator]
    def delete_rows(self,
                    search_criteria: RowInputData | str,
                    header_row: int = 1,
                    partial_match: bool = False,
                    first_only: bool = False) -> int:
        """Deletes all rows matching ``search_criteria``. Returns the count of deleted rows.

        See the `library introduction`_ for details on ``search_criteria`` and ``partial_match``.
        Streaming / read-only mode raises ``LibraryException``.

        Arguments:
        - ``search_criteria``: Filter identifying which rows to delete — see `library description`_ for format details.
        - ``header_row``: Row that contains the column headers (row 1 = first row). Defaults to ``1``.
        - ``partial_match``: Substring matching when ``True`` — see `library description`_. Defaults to ``False``.
        - ``first_only``: Delete only the first matching row when ``True``. Defaults to ``False``.

        Examples:
        | Load Workbook  | ${CURDIR}/data.xlsx |                                    |                 |
        | ${count} =     | Delete Rows         | ${{{'Product ID': 'P-001'}}}       |                 |
        | Should Be Equal As Integers | ${count} | 1 |                            |
        | Delete Rows    | Location=Online     | partial_match=True | first_only=True |
        | Save Workbook  |                     |                                    |                 |
        """
        if self._active_workbook:
            return self._active_workbook.delete_rows(
                search_criteria=search_criteria,
                header_row=header_row,
                partial_match=partial_match,
                first_only=first_only,
            )
        return 0

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