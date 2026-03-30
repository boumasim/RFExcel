from rfexcel.RFExcelLibrary import RFExcelLibrary
from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
XLSX_EDIT = "xlsx_edit"
XLSX_STREAM = "xlsx_stream"
CSV_EDIT = "csv_edit"
CSV_STREAM = "csv_stream"
XLS_EDIT = "xls_edit"
XLS_ON_DEMAND = "xls_on_demand"

XLSX_FORMAT = "xlsx"
CSV_FORMAT = "csv"
XLS_FORMAT = "xls"
FORMAT_LIST = [XLSX_FORMAT, CSV_FORMAT, XLS_FORMAT]
EDITABLE_FORMAT_LIST = [XLSX_FORMAT, CSV_FORMAT]

# ---------------------------------------------------------------------------
# Backend registry
# ---------------------------------------------------------------------------

BACKENDS: dict[str, tuple[str, bool]] = {
    XLSX_EDIT:     (XLSX_FILE, False),
    XLSX_STREAM:   (XLSX_FILE, True),
    CSV_EDIT:      (CSV_FILE,  False),
    CSV_STREAM:    (CSV_FILE,  True),
    XLS_EDIT:      (XLS_FILE,  False),
    XLS_ON_DEMAND: (XLS_FILE,  True),
}

FORMAT_FILE: dict[str, str] = {
    XLSX_FORMAT: XLSX_FILE,
    CSV_FORMAT:  CSV_FILE,
    XLS_FORMAT:  XLS_FILE,
}


def open_backend(lib: RFExcelLibrary, backend_name: str) -> None:
    path, read_only = BACKENDS[backend_name]
    lib.load_workbook(path, read_only=read_only)

BACKEND_NAMES = list(BACKENDS)

# ---------------------------------------------------------------------------
# Common data
# ---------------------------------------------------------------------------
XLSX_HEADERS = ["Product ID", "Description", "Price", "Location"]
XLS_HEADERS  = ["Index", "First Name", "Last Name", "Gender", "Country", "Age"]
CSV_HEADERS = ["Product ID", "Description", "Price", "Location"]

XLSX_ROWS = [
    {"Product ID": "P-200", "Description": "Wireless Mouse",            "Price": 25.5,  "Location": "Warehouse A, Shelf 2"},
    {"Product ID": "P-201", "Description": "Keyboard, Mechanical",      "Price": 89.99, "Location": "Store Front"},
    {"Product ID": "P-202", "Description": "Monitor 24-inch",           "Price": 150,   "Location": "Paris, France"},
    {"Product ID": "P-203", "Description": "USB Cable",                 "Price": 5.99,  "Location": "OnlineP"},
]

CSV_ROWS = [
    {"Product ID": "P-200", "Description": "Wireless Mouse",            "Price": 25.5,  "Location": "Warehouse A, Shelf 2"},
    {"Product ID": "P-201", "Description": "Keyboard, Mechanical, RGB", "Price": 89.99, "Location": "Store Front"},
    {"Product ID": "P-202", "Description": "Monitor 24-inch",           "Price": 150,   "Location": "Paris, France"},
    {"Product ID": "P-203", "Description": "USB Cable, 3ft",            "Price": 5.99,  "Location": "Online"},
]

XLS_ROWS = [
    {"Index": 1, "First Name": "Dulce",    "Last Name": "Abril",     "Gender": "Female", "Country": "United States", "Age": 32},
    {"Index": 2, "First Name": "Mara",     "Last Name": "Hashimoto", "Gender": "Female", "Country": "Great Britain", "Age": 25},
    {"Index": 3, "First Name": "Philip",   "Last Name": "Gent",      "Gender": "Male",   "Country": "France",        "Age": 36},
    {"Index": 4, "First Name": "Kathleen", "Last Name": "Hanner",    "Gender": "Female", "Country": "United States", "Age": 25},
    {"Index": 5, "First Name": "Nereida",  "Last Name": "Magwood",   "Gender": "Female", "Country": "United States", "Age": 58},
    {"Index": 6, "First Name": "Gaston",   "Last Name": "Brumm",     "Gender": "Male",   "Country": "United States", "Age": 24},
    {"Index": 7, "First Name": "Etta",     "Last Name": "Hurn",      "Gender": "Female", "Country": "Great Britain", "Age": 56},
    {"Index": 8, "First Name": "Earlean",  "Last Name": "Melgar",    "Gender": "Female", "Country": "United States", "Age": 27},
    {"Index": 9, "First Name": "Vincenza", "Last Name": "Weiland",   "Gender": "Female", "Country": "United States", "Age": 40},
]
