XLSX_SUFFIX = ".xlsx"
XLS_SUFFIX = ".xls"
CSV_SUFFIX = '.csv'

VALID_SUFFIXES = {XLSX_SUFFIX, XLS_SUFFIX, CSV_SUFFIX}

BASE_DIALECT = 'excel'
BASE_ENCODING = 'utf-8'

#error messages
CSV_NOT_SUPPORTED_MSG = "Operation is not supported for CSV files"
XLSX_NOT_SUPPORTED_MSG = "Operation is not supported for XLSX files"
XLS_NOT_SUPPORTED_MSG = "Operation is not supported for XLS files"