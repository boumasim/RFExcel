from typing import Any

from tests.pyth.test_data import (CSV_EDIT, CSV_STREAM, XLS_EDIT,
                                  XLS_ON_DEMAND, XLSX_EDIT, XLSX_STREAM)

RowData = dict[str, Any]
RowsData = list[RowData]

FULL_ROW_BY_BACKEND: dict[str, RowData] = {
    XLSX_EDIT: {"Product ID": "P-999", "Description": "Widget", "Price": 9.99, "Location": "Online"},
    XLSX_STREAM: {"Product ID": "P-999", "Description": "Widget", "Price": 9.99, "Location": "Online"},
    CSV_EDIT: {"Product ID": "P-999", "Description": "Widget", "Price": 9.99, "Location": "Online"},
    CSV_STREAM: {"Product ID": "P-999", "Description": "Widget", "Price": 9.99, "Location": "Online"},
    XLS_EDIT: {"First Name": "Jane", "Last Name": "Doe"},
    XLS_ON_DEMAND: {"First Name": "Jane", "Last Name": "Doe"},
}
EXPECTED_FULL_ROW_BY_BACKEND: dict[str, RowData] = {
    XLSX_EDIT: {"Product ID": "P-999", "Description": "Widget", "Price": 9.99, "Location": "Online"},
    XLSX_STREAM: {"Product ID": "P-999", "Description": "Widget", "Price": 9.99, "Location": "Online"},
    CSV_EDIT: {"Product ID": "P-999", "Description": "Widget", "Price": 9.99, "Location": "Online"},
    CSV_STREAM: {"Product ID": "P-999", "Description": "Widget", "Price": 9.99, "Location": "Online"},
    XLS_EDIT: {"Index": "", "First Name": "Jane", "Last Name": "Doe", "Gender": "", "Country": "", "Age": ""},
    XLS_ON_DEMAND: {"Index": "", "First Name": "Jane", "Last Name": "Doe", "Gender": "", "Country": "", "Age": ""},
}


PARTIAL_ROW_BY_BACKEND: dict[str, RowData] = {
    XLSX_EDIT: {"Product ID": "P-888", "Price": 0.01},
    XLSX_STREAM: {"Product ID": "P-888", "Price": 0.01},
    CSV_EDIT: {"Description": "Only Desc"},
    CSV_STREAM: {"Description": "Only Desc"},
    XLS_EDIT: {"First Name": "Only"},
    XLS_ON_DEMAND: {"First Name": "Only"},
}
EXPECTED_PARTIAL_ROW_BY_BACKEND: dict[str, RowData] = {
    XLSX_EDIT: {"Product ID": "P-888", "Description": "", "Price": 0.01, "Location": ""},
    XLSX_STREAM: {"Product ID": "P-888", "Description": "", "Price": 0.01, "Location": ""},
    CSV_EDIT: {"Product ID": "", "Description": "Only Desc", "Price": "", "Location": ""},
    CSV_STREAM: {"Product ID": "", "Description": "Only Desc", "Price": "", "Location": ""},
    XLS_EDIT: {"Index": "", "First Name": "Only", "Last Name": "", "Gender": "", "Country": "", "Age": ""},
    XLS_ON_DEMAND: {"Index": "", "First Name": "Only", "Last Name": "", "Gender": "", "Country": "", "Age": ""},
}


UNKNOWN_KEY_ROW_BY_BACKEND: dict[str, RowData] = {
    XLSX_EDIT: {"Product ID": "P-777", "NonExistent": "ignored"},
    XLSX_STREAM: {"Product ID": "P-777", "NonExistent": "ignored"},
    CSV_EDIT: {"Product ID": "P-777", "NonExistent": "ignored"},
    CSV_STREAM: {"Product ID": "P-777", "NonExistent": "ignored"},
    XLS_EDIT: {"First Name": "Ignored", "NonExistent": "ignored"},
    XLS_ON_DEMAND: {"First Name": "Ignored", "NonExistent": "ignored"},
}
EXPECTED_UNKNOWN_KEY_ROW_BY_BACKEND: dict[str, RowData] = {
    XLSX_EDIT: {"Product ID": "P-777", "Description": "", "Price": "", "Location": ""},
    XLSX_STREAM: {"Product ID": "P-777", "Description": "", "Price": "", "Location": ""},
    CSV_EDIT: {"Product ID": "P-777", "Description": "", "Price": "", "Location": ""},
    CSV_STREAM: {"Product ID": "P-777", "Description": "", "Price": "", "Location": ""},
    XLS_EDIT: {"Index": "", "First Name": "Ignored", "Last Name": "", "Gender": "", "Country": "", "Age": ""},
    XLS_ON_DEMAND: {"Index": "", "First Name": "Ignored", "Last Name": "", "Gender": "", "Country": "", "Age": ""},
}


ORDERED_ROWS_BY_BACKEND: dict[str, tuple[RowData, RowData]] = {
    XLSX_EDIT: (
        {"Product ID": "P-A", "Description": "Alpha", "Price": 1.0, "Location": "Shelf-A"},
        {"Product ID": "P-B", "Description": "Beta", "Price": 2.0, "Location": "Shelf-B"},
    ),
    XLSX_STREAM: (
        {"Product ID": "P-A", "Description": "Alpha", "Price": 1.0, "Location": "Shelf-A"},
        {"Product ID": "P-B", "Description": "Beta", "Price": 2.0, "Location": "Shelf-B"},
    ),
    CSV_EDIT: (
        {"Product ID": "P-A", "Description": "Alpha", "Price": 1.0, "Location": "Shelf-A"},
        {"Product ID": "P-B", "Description": "Beta", "Price": 2.0, "Location": "Shelf-B"},
    ),
    CSV_STREAM: (
        {"Product ID": "P-A", "Description": "Alpha", "Price": 1.0, "Location": "Shelf-A"},
        {"Product ID": "P-B", "Description": "Beta", "Price": 2.0, "Location": "Shelf-B"},
    ),
    XLS_EDIT: (
        {"First Name": "Alice", "Last Name": "Smith"},
        {"First Name": "Bob", "Last Name": "Jones"},
    ),
    XLS_ON_DEMAND: (
        {"First Name": "Alice", "Last Name": "Smith"},
        {"First Name": "Bob", "Last Name": "Jones"},
    ),
}
EXPECTED_ORDERED_ROWS_BY_BACKEND: dict[str, tuple[RowData, RowData]] = {
    XLSX_EDIT: (
        {"Product ID": "P-A", "Description": "Alpha", "Price": 1.0, "Location": "Shelf-A"},
        {"Product ID": "P-B", "Description": "Beta", "Price": 2.0, "Location": "Shelf-B"},
    ),
    XLSX_STREAM: (
        {"Product ID": "P-A", "Description": "Alpha", "Price": 1.0, "Location": "Shelf-A"},
        {"Product ID": "P-B", "Description": "Beta", "Price": 2.0, "Location": "Shelf-B"},
    ),
    CSV_EDIT: (
        {"Product ID": "P-A", "Description": "Alpha", "Price": 1.0, "Location": "Shelf-A"},
        {"Product ID": "P-B", "Description": "Beta", "Price": 2.0, "Location": "Shelf-B"},
    ),
    CSV_STREAM: (
        {"Product ID": "P-A", "Description": "Alpha", "Price": 1.0, "Location": "Shelf-A"},
        {"Product ID": "P-B", "Description": "Beta", "Price": 2.0, "Location": "Shelf-B"},
    ),
    XLS_EDIT: (
        {"Index": "", "First Name": "Alice", "Last Name": "Smith", "Gender": "", "Country": "", "Age": ""},
        {"Index": "", "First Name": "Bob", "Last Name": "Jones", "Gender": "", "Country": "", "Age": ""},
    ),
    XLS_ON_DEMAND: (
        {"Index": "", "First Name": "Alice", "Last Name": "Smith", "Gender": "", "Country": "", "Age": ""},
        {"Index": "", "First Name": "Bob", "Last Name": "Jones", "Gender": "", "Country": "", "Age": ""},
    ),
}


ADD_ROWS_PARTIAL_INPUT_BY_BACKEND: dict[str, RowsData] = {
    XLSX_EDIT: [{"Product ID": "P-010"}, {"Price": 5.0}],
    XLSX_STREAM: [{"Product ID": "P-010"}, {"Price": 5.0}],
    CSV_EDIT: [{"Product ID": "P-010"}, {"Price": 5.0}],
    CSV_STREAM: [{"Product ID": "P-010"}, {"Price": 5.0}],
    XLS_EDIT: [{"First Name": "OnlyOne"}, {"Age": 44}],
    XLS_ON_DEMAND: [{"First Name": "OnlyOne"}, {"Age": 44}],
}
EXPECTED_ADD_ROWS_PARTIAL_BY_BACKEND: dict[str, tuple[RowData, RowData]] = {
    XLSX_EDIT: (
        {"Product ID": "P-010", "Description": "", "Price": "", "Location": ""},
        {"Product ID": "", "Description": "", "Price": 5.0, "Location": ""},
    ),
    XLSX_STREAM: (
        {"Product ID": "P-010", "Description": "", "Price": "", "Location": ""},
        {"Product ID": "", "Description": "", "Price": 5.0, "Location": ""},
    ),
    CSV_EDIT: (
        {"Product ID": "P-010", "Description": "", "Price": "", "Location": ""},
        {"Product ID": "", "Description": "", "Price": 5.0, "Location": ""},
    ),
    CSV_STREAM: (
        {"Product ID": "P-010", "Description": "", "Price": "", "Location": ""},
        {"Product ID": "", "Description": "", "Price": 5.0, "Location": ""},
    ),
    XLS_EDIT: (
        {"Index": "", "First Name": "OnlyOne", "Last Name": "", "Gender": "", "Country": "", "Age": ""},
        {"Index": "", "First Name": "", "Last Name": "", "Gender": "", "Country": "", "Age": 44},
    ),
    XLS_ON_DEMAND: (
        {"Index": "", "First Name": "OnlyOne", "Last Name": "", "Gender": "", "Country": "", "Age": ""},
        {"Index": "", "First Name": "", "Last Name": "", "Gender": "", "Country": "", "Age": 44},
    ),
}
