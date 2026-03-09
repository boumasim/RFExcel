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

Get Rows with Search Criteria
    [Documentation]     Get rows matching search criteria and verify results
    Load Workbook    path=${RESOURCES}/data.csv      read_only=true
    ${criteria}=    Create Dictionary    Product ID=P-201    Price=89.99
    ${matching_rows}=    Get Rows    search_criteria=${criteria}
    Log    Matching Rows: ${matching_rows}

Sheet test
    [Documentation]     Test sheet operations: create, switch, and verify data
    Load Workbook    path=${RESOURCES}/test.xlsx
    ${sheets}=     List Sheet Names
    Add Sheet       name=Test 1
    ${sheet1}=     List Sheet Names
    Switch Sheet    name=List 1
    Delete Sheet    name=Test1
    ${sheet_names}=     List Sheet Names
    Log    Remaining Sheets: ${sheet_names}

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