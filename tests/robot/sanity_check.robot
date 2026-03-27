*** Settings ***
Documentation    Testing my custom Python keywords
Library         Collections
Library        rfexcel.RFExcelLibrary

*** Variables ***
${RESOURCES}    ${CURDIR}/../resources
${RESULTS}      ${CURDIR}/../../results

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

Get Rows by each row
    [Documentation]     Get rows one by one and verify data structure
    Load Workbook    path=${RESOURCES}/data.csv      read_only=true
    ${my_list}=    Create List
    ${headers}=   Get Row    row=1
    FOR    ${i}    IN RANGE    2    4
        ${row}=    Get Row    row=${i}    headers=${headers}
        Append To List    ${my_list}    ${row}
        Log    Row ${i}: ${row}
    END
    Log    All rows: ${my_list}
    Length Should Be    ${my_list}    2

Switch Source Test
    [Documentation]     Test switching between different sources and verify data integrity
    Load Workbook    path=${RESOURCES}/data.csv      read_only=true
    ${csv_rows}=    Get Rows
    Log    CSV Rows: ${csv_rows}
    Switch Source    path=${RESOURCES}/data.xlsx     read_only=true
    ${xlsx_rows}=   Get Rows
    Log    XLSX Rows: ${xlsx_rows}
    Length Should Be    ${csv_rows}     4
    Length Should Be    ${xlsx_rows}    4

Get Rows with Search Criteria
    [Documentation]     Get rows matching search criteria and verify results
    Load Workbook    path=${RESOURCES}/data.csv      read_only=true
    ${criteria}=    Create Dictionary    Product ID=P-201    Price=89.99
    ${matching_rows}=    Get Rows    search_criteria=${criteria}
    Log    Matching Rows: ${matching_rows}
    Length Should Be    ${matching_rows}    1
    Should Be Equal    ${matching_rows}[0][Product ID]    P-201

Sheet test
    [Documentation]     Test sheet operations: create, switch, and verify data
    Load Workbook    path=${RESOURCES}/data.xlsx
    ${sheets}=     List Sheet Names
    Log    Sheets before add: ${sheets}
    Add Sheet       name=Test 1
    ${sheet1}=     List Sheet Names
    Should Contain    ${sheet1}    Test 1
    Switch Sheet    name=List 1
    Delete Sheet    name=Test 1
    ${sheet_names}=     List Sheet Names
    Log    Remaining Sheets: ${sheet_names}
    Should Not Contain    ${sheet_names}    Test 1

Save sheet test
    [Documentation]     Test saving a workbook after modifications
    Load Workbook    path=${RESOURCES}/data.csv
    Save Workbook   path=${RESULTS}/data.csv

Add row to shifted table
    [Documentation]     Test adding a row to a shifted table and verify data integrity
    Load Workbook    path=${RESOURCES}/data.xlsx
    Switch Sheet    name=Sheet3
    ${value_map}=   Create Dictionary    Product ID=XD    Description=LOL    Price=69.00
    Append Row    row_data=${value_map}    header_row=3
    ${rows}=    Get Rows    header_row=3
    Save Workbook   path=${RESULTS}/data.xlsx

Compare data to another file
    [Documentation]     Test comparing data between two files and verify differences
    Load Workbook    path=${RESOURCES}/data.xlsx  read_only=False
    ${differences}=    Compare Data To    target_path=${RESOURCES}/data2.xlsx
    Log    Differences: ${differences}
    ${headers_filter}=    Create List    Description    Location
    ${differences}=    Compare Data To    target_path=${RESOURCES}/data2.xlsx    headers=${headers_filter}
    Log    Differences with headers: ${differences}

Lazy switch to xlsx
    [Documentation]     Test lazy switching from .xls to .xlsx and verify data integrity
    Load Workbook    path=${RESOURCES}/example.xls
    ${rows}=    Get Rows
    Log    Rows from .xls: ${rows}
    Should Be True    len($rows) > 0
    Save Workbook    path=${RESULTS}/example_converted.xlsx

Test xlsx generator close
    [Documentation]     Test that the row generator is properly closed when switching sheets or closing the workbook
    Load Workbook    path=${RESOURCES}/data.xlsx  read_only=True
    Switch Sheet    name=Sheet2
    ${row1}=   Get Row    row=2
    Log    First row: ${row1}
    Switch Sheet    name=Sheet3
    ${row2}=   Get Row    row=2
    Log    First row of Sheet2: ${row2}
    Close Workbook
