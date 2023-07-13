*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF.

Library             RPA.Browser.Selenium    # auto_close=${FALSE}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.PDF


*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open the Intranet Website
    # Maximize Browser Window
    Login
    Download the Excel File
    # Fill and submit the form
    # Fill and submit the form for one person
    Fill the form using the data from the Excel file
    Collect the Result
    Export The Table As Pdf
    [Teardown]    Logout and close the Browser


*** Keywords ***
# Minimal Tasks
#    Log

Open the Intranet Website
    Open Available Browser    https://robotsparebinindustries.com/    # args=--start-maximized

# Maximize Browser Window
#    Set Window Size    Maximize

Login
    Input Text    id:username    maria
    Input Password    id:password    thoushallnotpass
    Submit Form
    Wait Until Page Contains Element    id:sales-form

Download the Excel File
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=${True}

# Fill and submit the form
#    Input Text    id:firstname    Kishore
#    Input Text    id:lastname    Paul
#    Input Text    id:salesresult    50000
#    Select From List By Value    id:salestarget    10000
#    Click Button    Submit

Fill and submit the form for one person
    [Arguments]    ${sales_rep}
    Input Text    id:firstname    ${sales_rep}[First Name]
    Input Text    id:lastname    ${sales_rep}[Last Name]
    Input Text    id:salesresult    ${sales_rep}[Sales]
    Select From List By Value    id:salestarget    ${sales_rep}[Sales Target]
    Click Button    Submit

Fill the form using the data from the Excel file
    Open Workbook    SalesData.xlsx
    ${sales_reps}=    Read Worksheet As Table    header=${True}
    Close Workbook
    FOR    ${sales_rep}    IN    @{sales_reps}
        Fill and submit the form for one person    ${sales_rep}
    END
# Fill and submit the form for one person

Collect the Result
    Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}sales_summary.png

Export The Table As Pdf
    Wait Until Element Is Visible    id:sales-results
    ${sales_result_html}=    Get Element Attribute    id:sales-results    outerHTML
    Html To Pdf    ${sales_result_html}    ${OUTPUT_DIR}${/}sales_result.pdf

Logout and close the Browser
    Click Button    Log out
    Close Browser
