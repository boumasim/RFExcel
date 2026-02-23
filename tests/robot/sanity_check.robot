*** Settings ***
Documentation    Testing my custom Python keywords
Library         Collections
Library        rfexcel.RFExcelLibrary

*** Variables ***
${RESOURCES}    ${CURDIR}/../resources
${RESULTS}      ${CURDIR}/../../results

*** Test Cases ***
Load XLS File Read Only
    [Documentation]    Load XLS file in streaming (read-only) mode
    Load Workbook    path=${RESOURCES}/example.xls    read_only=True
    Print
    Close Workbook

Load XLS File Edit Mode
    [Documentation]    Load XLS file in edit mode (standard reader)
    Load Workbook    path=${RESOURCES}/example.xls    read_only=False
    Print
    Close Workbook

Load XLSX File Read Only
    [Documentation]    Load XLSX file in streaming (read-only) mode
    Load Workbook    path=${RESOURCES}/data.xlsx    read_only=True
    Print
    Close Workbook

Load XLSX File Edit Mode
    [Documentation]    Load XLSX file in edit mode
    Load Workbook    path=${RESOURCES}/data.xlsx    read_only=False
    Print
    Close Workbook

Load CSV File Read Only
    [Documentation]    Load CSV file in streaming (read-only) mode
    Load Workbook    path=${RESOURCES}/data.csv    read_only=True
    Print

Load CSV File Edit Mode
    [Documentation]    Load CSV file in edit mode (buffered in memory)
    Load Workbook    path=${RESOURCES}/data.csv    read_only=False
    Print
    Close Workbook

Create New XLSX File
    [Documentation]    Create a new XLSX file from scratch
    Create Workbook    path=${RESULTS}/test_created.xlsx
    Print
    Close Workbook

Create New CSV File
    [Documentation]    Create a new CSV file from scratch
    Create Workbook    path=${RESULTS}/test_created.csv
    Print
    Close Workbook

Get Rows from Workbook
    [Documentation]     Get rows from workbook and verify data structure
    Load Workbook    path=${RESOURCES}/data.csv      read_only=true
    ${rows}=    Get Rows
    Log    ${rows}

Get Rows by each row
    [Documentation]     Get rows one by one and verify data structure
    Load Workbook    path=${RESOURCES}/data.csv      read_only=true
    ${my_list}=    Create List
    ${headers}=   Get Row    row=2
    FOR    ${i}    IN RANGE    3    5
        ${row}=    Get Row    row=${i}    headers=${headers}
        Append To List    ${my_list}    ${row}
        Log    Row ${i}: ${row}
    END
    Log    All rows: ${my_list}

Switch Source Test
    [Documentation]     Test switching between different sources and verify data integrity
    Load Workbook    path=${RESOURCES}/data.csv      read_only=true
    ${csv_rows}=    Get Rows
    Log    CSV Rows: ${csv_rows}
    Switch Source    path=${RESOURCES}/data.xlsx     read_only=true
    ${xlsx_rows}=   Get Rows
    Log    XLSX Rows: ${xlsx_rows}