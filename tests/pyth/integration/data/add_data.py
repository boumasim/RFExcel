from tests.pyth.test_data import RowData, RowsData

FULL_ROW: RowData = {"Product ID": "P-999", "Description": "Widget", "Price": 9.99, "Location": "Online"}
EXPECTED_FULL_ROW: RowData = {"Product ID": "P-999", "Description": "Widget", "Price": 9.99, "Location": "Online"}

PARTIAL_ROW: RowData= {"Product ID": "P-888", "Price": 0.01}
EXPECTED_PARTIAL_ROW: RowData = {"Product ID": "P-888", "Description": "", "Price": 0.01, "Location": ""}

UNKNOWN_KEY_ROW: RowData = {"Product ID": "P-777", "NonExistent": "ignored"}
EXPECTED_UNKNOWN_KEY_ROW: RowData = {"Product ID": "P-777", "Description": "", "Price": "", "Location": ""}


ORDERED_ROWS: tuple[RowData, RowData] = (
        {"Product ID": "P-A", "Description": "Alpha", "Price": 1.0, "Location": "Shelf-A"},
        {"Product ID": "P-B", "Description": "Beta", "Price": 2.0, "Location": "Shelf-B"},
    )
EXPECTED_ORDERED_ROWS: tuple[RowData, RowData] = (
        {"Product ID": "P-A", "Description": "Alpha", "Price": 1.0, "Location": "Shelf-A"},
        {"Product ID": "P-B", "Description": "Beta", "Price": 2.0, "Location": "Shelf-B"},
    )

ADD_ROWS_PARTIAL_INPUT: RowsData = [{"Product ID": "P-010"}, {"Price": 5.0}]
EXPECTED_ADD_ROWS_PARTIAL: tuple[RowData, RowData] = (
        {"Product ID": "P-010", "Description": "", "Price": "", "Location": ""},
        {"Product ID": "", "Description": "", "Price": 5.0, "Location": ""},
    )
