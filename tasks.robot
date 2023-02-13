*** Settings ***
Documentation       Orders robots from RobotSpareBin Industries Inc.
...                 Saves the order HTML receipt as a PDF file.
...                 Saves the screenshot of the ordered robot.
...                 Embeds the screenshot of the robot to the PDF receipt.
...                 Creates ZIP archive of the receipts and the images.

Library             RPA.Browser.Selenium    auto_close=${True}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.Excel.Application
Library             RPA.Desktop.Windows
Library             RPA.Tables
Library             RPA.PDF
Library             RPA.Desktop
Library             RPA.Archive


*** Variables ***
${GLOBAL_RETRY_AMOUNT}=         5x
${GLOBAL_RETRY_INTERVAL}=       10s
${OUTPUT_DIR}=                  output
${URL}=                         https://robotsparebinindustries.com/#/robot-order


*** Tasks ***
Order robots from RobotSpareBin Industries Inc
    Download the Excel file
    Open the robot order website
    ${orders}=    Get Orders
    FOR    ${row}    IN    @{orders}
        Wait Until Element Is Visible    class:alert-buttons
        Click Button    OK
        Fill the form    ${row}
        Preview the robot
        Wait Until Keyword Succeeds    ${GLOBAL_RETRY_AMOUNT}    ${GLOBAL_RETRY_INTERVAL}    Submit the order
        ${pdf}=    Store the receipt as a PDF file    ${row}[Order number]
        ${screenshot}=    Take a screenshot of the robot    ${row}[Order number]
        Embed the robot screenshot to the receipt PDF file    ${screenshot}    ${pdf}
        Go to order another robot
    END
    Create ZIP package for receipts


*** Keywords ***
Download the Excel file
    Download    https://robotsparebinindustries.com/orders.csv    overwrite=True

Open the robot order website
    Open Available Browser    ${URL}

Get Orders
    ${orders}=    Read table from CSV    orders.csv    header=True
    RETURN    ${orders}

Fill the form
    [Arguments]    ${row}
    Wait Until Element Is Visible    id:head
    Select From List By Value    id:head    ${row}[Head]
    Select Radio Button    body    ${row}[Body]
    Input Text    id:address    ${row}[Address]
    Input Text    css:form .form-group:nth-child(3) input    ${row}[Legs]

Preview the robot
    Click Button    Preview

Submit the order
    Click Button    Order
    Wait Until Element Is Visible    id:receipt

Go to order another robot
    Click Button    Order another robot

Store the receipt as a PDF file
    [Arguments]    ${order_number}
    Wait Until Element Is Visible    id:receipt
    ${receipt_html}=    Get Element Attribute    id:receipt    outerHTML
    Html To Pdf    ${receipt_html}    ${OUTPUT_DIR}${/}${order_number}.pdf
    RETURN    ${OUTPUT_DIR}${/}${order_number}.pdf

Take a screenshot of the robot
    [Arguments]    ${order_number}
    Wait Until Element Is Visible    id:robot-preview-image
    Capture Element Screenshot    id:robot-preview-image    ${OUTPUT_DIR}${/}${order_number}.png
    RETURN    ${OUTPUT_DIR}${/}${order_number}.PNG

Embed the robot screenshot to the receipt PDF file
    [Arguments]    ${screenshot}    ${pdf}
    ${files}=    Create List    ${screenshot}
    Add Files To Pdf    ${files}    ${pdf}    append=True

Create ZIP package for receipts
    ${zip_file_name}=    Set Variable    ${OUTPUT_DIR}/Receipts.zip
    Archive Folder With Zip    ${OUTPUT_DIR}    ${zip_file_name}
