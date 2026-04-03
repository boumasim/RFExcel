from typing import cast

from tests.pyth.test_data import SHEET1_ROWS, RowData, RowsData

DELETE_ROW_SCENARIOS: tuple[tuple[str, int, RowsData], ...] = (
    ("first_data_row", 2, cast(RowsData, SHEET1_ROWS[1:])),
    (
        "second_data_row",
        3,
        cast(RowsData, [SHEET1_ROWS[0], *SHEET1_ROWS[2:]]),
    ),
    ("last_data_row", len(SHEET1_ROWS) + 1, cast(RowsData, SHEET1_ROWS[:-1])),
)

DELETE_ROWS_SINGLE_MATCH_CRITERIA: RowData = {"Product ID": "P-200"}
DELETE_ROWS_SECOND_MATCH_CRITERIA: RowData = {"Product ID": "P-201"}
DELETE_ROWS_NO_MATCH_CRITERIA: RowData = {"Product ID": "NONEXISTENT"}
DELETE_ROWS_PARTIAL_MATCH_CRITERIA: RowData = {"Location": "Warehouse"}
DELETE_ROWS_NUMERIC_MATCH_CRITERIA: RowData = {"Price": "150"}
DELETE_ROWS_DUPLICATE_MATCH_CRITERIA: RowData = {"Location": "SAME"}
DELETE_ROWS_DUPLICATE_UPDATE_CRITERIA: tuple[RowData, RowData] = (
    {"Product ID": "P-201"},
    {"Product ID": "P-202"},
)
DELETE_ROWS_DUPLICATE_UPDATE_VALUES: RowData = {"Location": "SAME"}

EXPECTED_ROWS_AFTER_DELETE_SINGLE_MATCH: RowsData = cast(RowsData, SHEET1_ROWS[1:])
EXPECTED_ROWS_AFTER_DELETE_NUMERIC_MATCH: RowsData = cast(
    RowsData,
    [SHEET1_ROWS[0], SHEET1_ROWS[1], SHEET1_ROWS[3]],
)
EXPECTED_PARTIAL_MATCH_DELETE_COUNT = 1
EXPECTED_DUPLICATE_DELETE_COUNT = 2
EXPECTED_DUPLICATE_ROWS_REMAINING_AFTER_ONE_ROW_DELETE = 1