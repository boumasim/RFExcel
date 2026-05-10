# RFExcel

A [Robot Framework](https://robotframework.org/) library for reading, writing, and manipulating Excel and CSV files.

[![PyPI version](https://img.shields.io/pypi/v/RFExcel.svg)](https://pypi.org/project/RFExcel/)
[![Python](https://img.shields.io/pypi/pyversions/RFExcel.svg)](https://pypi.org/project/RFExcel/)
[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](LICENSE)

## Overview

RFExcel provides Robot Framework keywords for verifying and manipulating spreadsheet data in `.xlsx`, `.xls`, and `.csv` formats. It preserves native Python types (`int`, `float`, `bool`, `datetime`) rather than coercing everything to strings, making data assertions precise and reliable.

For full keyword reference see [API documentation](https://boumasim.github.io/RFExcel/RFExcel.html).

Library mainly supports table like operations. For example of such a file, see tests/resources

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
Create New CSV File
    [Documentation]    Create a new CSV file from scratch
    Create Workbook    path=${RESULTS}/test_created.csv
    Close Workbook

Get Rows from Workbook
    [Documentation]     Get rows from workbook and verify data structure
    Load Workbook    path=${RESOURCES}/data.csv      read_only=true
    ${rows}=    Get Rows  one_row=true
    Log    ${rows}
    Should Not Be Empty    ${rows}
    Dictionary Should Contain Key    ${rows}    Product ID
```

More examples in /tests/robot/example.robot

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

> **Note:** Comparison for Date and Time values differ across formats, yet to be implemented

## Modes

**Edit mode** (`read_only=False`, default) — loads the full file into memory, supports reading and writing.

**Streaming mode** (`read_only=True`) — memory-efficient, read-only, strictly forward-only. Each row can be read only once, forward only reading supported.

## License

[Apache License 2.0](LICENSE)

## Others

<img src="https://fit.cvut.cz/static/images/fit-cvut-logo-en.svg" alt="FIT CTU logo" height="200">

This software was developed with the support of the **Faculty of Information Technology, Czech Technical University in Prague**.
For more information, visit [fit.cvut.cz](https://fit.cvut.cz).
