# RFExcel

A [Robot Framework](https://robotframework.org/) library for reading, writing, and manipulating Excel and CSV files.

[![PyPI version](https://img.shields.io/pypi/v/RFExcel.svg)](https://pypi.org/project/RFExcel/)
[![Python](https://img.shields.io/pypi/pyversions/RFExcel.svg)](https://pypi.org/project/RFExcel/)
[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](LICENSE)

## Overview

RFExcel provides Robot Framework keywords for verifying and manipulating spreadsheet data in `.xlsx`, `.xls`, and `.csv` formats. It preserves native Python types (`int`, `float`, `bool`, `datetime`) rather than coercing everything to strings, making data assertions precise and reliable.

## Features

- **Multi-format support** — `.xlsx`, `.xls` (read/write), and `.csv`
- **Dual-mode operation** — Edit mode (full in-memory read/write) and Streaming mode (memory-efficient, read-only, forward-only)
- **Native type preservation** — cell values returned as `int`, `float`, `bool`, `datetime`, or `str`
- **Flexible row filtering** — search criteria with exact or partial matching, AND logic across multiple columns
- **Sheet management** — switch, add, delete, and list sheets
- **Cell-level access** — get and set individual cells by coordinate (e.g. `A1`, `C3`)
- **Lazy XLS conversion** — `.xls` write operations convert in-memory to `.xlsx` without modifying the original file
- **Comparison** — diff two workbooks side-by-side with `Compare Data To`

## Supported Formats

| Format  | Edit mode | Streaming mode | Notes |
|---------|-----------|----------------|-------|
| `.xlsx` | yes       | yes            | Full read/write via openpyxl |
| `.xls`  | yes*      | yes*           | *Write triggers lazy in-memory conversion to `.xlsx`; original file unchanged |
| `.csv`  | yes       | yes            | No sheet concept; sheet keywords raise `OperationNotSupportedForFormat` |

## Installation

```shell
pip install RFExcel
```

**Dependencies:** `openpyxl>=3.0.0`, `robotframework>=7.0.0`, `xlrd>=2.0.2`

## Quick Start

```robotframework
*** Settings ***
Library    rfexcel.RFExcelLibrary

*** Test Cases ***
Read Rows From Excel
    Load Workbook    path=data.xlsx    read_only=True
    ${rows}=    Get Rows
    Length Should Be    ${rows}    4
    Should Be Equal    ${rows}[0][Product ID]    P-200

Filter Rows By Criteria
    Load Workbook    path=data.xlsx    read_only=True
    ${matching}=    Get Rows    search_criteria=Product ID=P-201;Price=89.99
    Length Should Be    ${matching}    1

Append And Save
    Load Workbook    path=data.xlsx
    ${row}=    Create Dictionary    Product ID=P-999    Description=New Item    Price=9.99
    Append Row    row_data=${row}
    Save Workbook    path=output.xlsx

Compare Two Files
    Load Workbook    path=baseline.xlsx
    ${differences}=    Compare Data To    target_path=updated.xlsx
    Should Be Empty    ${differences}
```

## Keywords

| Keyword | Description |
|---------|-------------|
| `Load Workbook` | Open an existing file in edit or streaming mode |
| `Create Workbook` | Create a new empty workbook |
| `Save Workbook` | Persist changes to disk |
| `Close Workbook` | Close the active workbook (done automatically at test end) |
| `Switch Source` | Swap the active file without closing first |
| `Get Rows` | Return all (or filtered) rows as a list of dicts |
| `Get Row` | Return a single row by row number |
| `Append Row` | Add a row at the end of the sheet |
| `Append Rows` | Add multiple rows at once |
| `Insert Row` | Insert a row at a specific position |
| `Delete Row` | Delete a row by its row number |
| `Delete Rows` | Delete all rows matching a search criterion |
| `Update Values` | Update column values in all matching rows |
| `Get Cell` | Get a single cell value by coordinate (e.g. `A1`) |
| `Set Cell` | Set a single cell value by coordinate |
| `Compare Data To` | Diff the active workbook against another file |
| `Switch Sheet` | Set the active sheet by name |
| `List Sheet Names` | Return all sheet names |
| `Add Sheet` | Add a new sheet |
| `Delete Sheet` | Remove a sheet by name |

For the full keyword reference and argument details see the [API documentation](https://boumasim.github.io/RFExcel/RFExcel.html).

## Search Criteria

Keywords that filter rows accept a `search_criteria` argument as either a `dict` or a semicolon-separated `key=value` string. Matching uses AND logic across all pairs.

```robotframework
# Dict syntax
${rows}=    Get Rows    search_criteria=${{ {"Product ID": "P-200", "Price": "25.5"} }}

# String syntax
${rows}=    Get Rows    search_criteria=Product ID=P-200;Price=25.5

# Partial matching
${rows}=    Get Rows    search_criteria=Description=Keyboard    partial_match=True
```

## Data Types

Cell values are returned as native Python types:

| Excel / CSV format  | Robot Framework type |
|---------------------|---------------------|
| Text / General      | `str` |
| Whole number        | `int` |
| Decimal / Currency  | `float` |
| Date / Time         | `datetime` |
| Boolean             | `bool` |
| Empty cell          | `""` |

> **Note:** Because types are preserved, use type-aware assertions — e.g. `Should Be Equal As Numbers` for numeric cells rather than plain `Should Be Equal`.

## Modes

**Edit mode** (`read_only=False`, default) — loads the full file into memory, supports reading and writing.

**Streaming mode** (`read_only=True`) — memory-efficient, read-only, strictly forward-only. Calling a read keyword twice on the same open workbook raises `StreamingViolationException`.

## License

[Apache License 2.0](LICENSE)
