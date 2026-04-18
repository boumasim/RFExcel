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
    ${matching_rows}=    Get Rows    search_criteria=Product ID=P-201;Price=89.99
    Log    Matching Rows: ${matching_rows}
    Length Should Be    ${matching_rows}    1
    Should Be Equal    ${matching_rows}[0][Product ID]    P-201

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
    Should Not Be Empty    ${rows}
    Save Workbook    path=${RESULTS}/example_converted.xlsx

Get Cell From Workbook
    [Documentation]     Get a single cell value by coordinate
    Load Workbook    path=${RESOURCES}/data.xlsx    read_only=True
    ${value}=    Get Cell    A1
    Should Be Equal    ${value}    Product ID
    ${price}=    Get Cell    C2
    Should Be Equal As Numbers    ${price}    25.5

Set Cell In Workbook
    [Documentation]     Set a single cell value by coordinate and verify
    Load Workbook    path=${RESOURCES}/data.xlsx
    Set Cell    A1    Updated Header
    ${value}=    Get Cell    A1
    Should Be Equal    ${value}    Updated Header
    Save Workbook    path=${RESULTS}/set_cell_result.xlsx

List Sheet Names From Workbook
    [Documentation]     List all sheet names from a workbook
    Load Workbook    path=${RESOURCES}/data.xlsx    read_only=True
    ${sheets}=    List Sheet Names
    Should Contain    ${sheets}    Sheet1
    Should Contain    ${sheets}    Sheet2
    Length Should Be    ${sheets}    3

Add Sheet To Workbook
    [Documentation]     Add a new sheet and verify it appears in the sheet list
    Load Workbook    path=${RESOURCES}/data.xlsx
    Add Sheet    SanitySheet
    ${sheets}=    List Sheet Names
    Should Contain    ${sheets}    SanitySheet
    Save Workbook    path=${RESULTS}/add_sheet_result.xlsx

Delete Sheet From Workbook
    [Documentation]     Delete a sheet and verify it is removed from the sheet list
    Load Workbook    path=${RESOURCES}/data.xlsx
    Delete Sheet    Sheet2
    ${sheets}=    List Sheet Names
    Should Not Contain    ${sheets}    Sheet2
    Save Workbook    path=${RESULTS}/delete_sheet_result.xlsx

Append Rows To Workbook
    [Documentation]     Append multiple rows at once and verify count increases
    Load Workbook    path=${RESOURCES}/data.xlsx
    ${rows_before}=    Get Rows
    ${row1}=    Create Dictionary    Product ID=P-901    Description=Widget A    Price=1.99    Location=Online
    ${row2}=    Create Dictionary    Product ID=P-902    Description=Widget B    Price=2.99    Location=Warehouse
    ${new_rows}=    Create List    ${row1}    ${row2}
    Append Rows    ${new_rows}
    ${rows_after}=    Get Rows
    Length Should Be    ${rows_after}    ${{ len(${rows_before}) + 2 }}
    Save Workbook    path=${RESULTS}/append_rows_result.xlsx

Insert Row Into Workbook
    [Documentation]     Insert a row at a specific position and verify it appears there
    Load Workbook    path=${RESOURCES}/data.xlsx
    ${new_row}=    Create Dictionary    Product ID=P-000    Description=First Item    Price=0.01    Location=Top
    Insert Row    row_data=${new_row}    row=2
    ${inserted}=    Get Rows    search_criteria=Product ID=P-000
    Length Should Be    ${inserted}    1
    Save Workbook    path=${RESULTS}/insert_row_result.xlsx

Update Values In Workbook
    [Documentation]     Update cells in matching rows and verify the count returned
    Load Workbook    path=${RESOURCES}/data.xlsx
    ${count}=    Update Values
    ...    search_criteria=Product ID=P-200
    ...    values=${{{"Price": 99.99}}}
    Should Be Equal As Integers    ${count}    1
    ${updated}=    Get Rows    search_criteria=Product ID=P-200
    Should Be Equal As Numbers    ${updated}[0][Price]    99.99
    Save Workbook    path=${RESULTS}/update_values_result.xlsx

Delete Rows From Workbook
    [Documentation]     Delete all rows matching a search criterion and verify the count
    Load Workbook    path=${RESOURCES}/data.xlsx
    ${count}=    Delete Rows    search_criteria=Product ID=P-201
    Should Be Equal As Integers    ${count}    1
    ${remaining}=    Get Rows    search_criteria=Product ID=P-201
    Length Should Be    ${remaining}    0
    Save Workbook    path=${RESULTS}/delete_rows_result.xlsx

Delete Row By Number
    [Documentation]     Delete a row by its row number and verify row count decreases
    Load Workbook    path=${RESOURCES}/data.xlsx
    ${rows_before}=    Get Rows
    Delete Row    2
    ${rows_after}=    Get Rows
    Length Should Be    ${rows_after}    ${{ len(${rows_before}) - 1 }}
    Save Workbook    path=${RESULTS}/delete_row_result.xlsx
