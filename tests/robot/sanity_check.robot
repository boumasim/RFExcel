*** Settings ***
Documentation    Testing my custom Python keywords
Library        rfexcel.RFExcelLibrary

*** Test Cases ***
Verify Workbook Creation
    [Documentation]    Checks if we can create a workbook without crashing

    Create Workbook   path=/home/bouma1/PycharmProjects/RFExcel/resources/xddd.xlsx
    Print