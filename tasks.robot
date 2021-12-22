*** Settings ***
Documentation     Orders robots from RobotSpareBin Industries Inc.
Library           RPA.Browser.Selenium    auto_close=${FALSE}
Library           RPA.Excel.Files
Library           RPA.HTTP
Library           RPA.PDF
Library           RPA.Desktop
Library           RPA.Tables
Library           RPA.Archive
Library           RPA.FileSystem
Library           RPA.Dialogs
Library           RPA.Robocloud.Secrets

*** Tasks ***
Order robots from RobotSpareBin Industries Inc
    Download the Excel file
    Open the robot order website
    Fill the form using the data from the Excel file
    [Teardown]    Close the Browser
    Create zip folder for PDFs

*** Keywords ***
Download the Excel file
    Add heading    Enter URL to Download Order CSV
    Add text input    Download_URL    label=Download URL    placeholder=Enter URL Here    rows=1
    ${result}=    Run dialog
    ${PopUpButton}    Set Variable    ${result}[Download_URL]
    Download    ${result}[Download_URL]    overwrite=True
# https://robotsparebinindustries.com/orders.csv

Get URL from Vault
    ${secret}=    RPA.Robocloud.Secrets.Get Secret    WebsiteDetails

Open the robot order website
    ${secret}=    RPA.Robocloud.Secrets.Get Secret    WebsiteDetails
    Open Available Browser    ${secret}[URL]
    Maximize Browser Window
    Sleep    2

 Close the annoying modal
    Wait Until Page Contains Element    class:btn.btn-dark
    Click Button    class:btn.btn-dark
    Sleep    2

 Fill the form for one order
    [Arguments]    ${sales_rep}
    Wait Until Page Contains Element    head
    Select From List By Index    head    ${sales_rep}[Head]
    Sleep    2
    Click Button    id:id-body-${sales_rep}[Body]
    Sleep    2
    Input Text    class:form-control    ${sales_rep}[Legs]
    Sleep    2
    Input Text    address    ${sales_rep}[Address]
    Sleep    2
    Click Button    preview
    Sleep    5
    Wait Until Page Contains Element    order
    Click Button    order
    Sleep    5
    ${boolOrder} =    Is Element Enabled    order
    IF    ${boolOrder}
        Click Button    order
        Sleep    5
    END
    ${boolOrder} =    Is Element Enabled    order
    IF    ${boolOrder}
        Click Button    order
        Sleep    5
        ${boolOrder} =    Is Element Enabled    order
    END
    ${boolOrder} =    Is Element Enabled    order
    IF    ${boolOrder}
        Click Button    order
        Sleep    5
    END
    ${boolOrder} =    Is Element Enabled    order
    IF    ${boolOrder}
        Click Button    order
        Sleep    5
    END

Store the receipt as pdf file
    [Arguments]    ${order_number}
    ${boolPreview} =    Is Element Visible    preview
    IF    ${boolPreview}
        Click Button    preview
        Sleep    5
        Click Button    order
        Sleep    5
    END
    Wait Until Page Contains Element    receipt    30
    ${receipt_html}=    Get Element Attribute    id:receipt    outerHTML
    Html To Pdf    ${receipt_html}    ${OUTPUT_DIR}${/}PDFs${/}order${order_number}.pdf    overwrite=True

 Take a screenshot of the robot
    [Arguments]    ${order_number}
    Screenshot    id:robot-preview-image    ${OUTPUT_DIR}${/}PDFs${/}order${order_number}.png

Embed the robot screenshot to the receipt PDF file
    [Arguments]    ${PDF_Path}    ${Screenshot_Path}
    ${Files_List}=    Create List
    ...    ${PDF_Path}
    ...    ${Screenshot_Path}
    Open Pdf    ${PDF_Path}
    Add Files To Pdf    ${Files_List}    ${PDF_Path}
    Close Pdf    ${PDF_Path}

Go to order another robot
    Wait Until Page Contains Element    order-another
    Click Button    order-another

Fill the form using the data from the Excel file
    ${orders_rep}=    Read table from CSV    orders.csv
    FOR    ${orders_rep}    IN    @{orders_rep}
        Close the annoying modal
        Fill the form for one order    ${orders_rep}
        Store the receipt as pdf file    ${orders_rep}[Order number]
        Take a screenshot of the robot    ${orders_rep}[Order number]
        Embed the robot screenshot to the receipt PDF file    ${OUTPUT_DIR}${/}PDFs${/}order${orders_rep}[Order number].pdf    ${OUTPUT_DIR}${/}PDFs${/}order${orders_rep}[Order number].png
        Go to order another robot
    END

Create zip folder for PDFs
    Archive Folder With Zip    ${OUTPUT_DIR}${/}PDFs    ${OUTPUT_DIR}${/}PDFs.zip

Close the Browser
    Close Browser
