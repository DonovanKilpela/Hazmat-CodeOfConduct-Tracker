import xlwings as xw
import time
import requests
from tabulate import tabulate

def findOpenBook():
    #for apps first
    app = xw.apps.active

    # Check if the workbook is already open
    openExcel = xw.Book('Hazmat_tracker.xlsm')

    # If the workbook is not found, you can handle the situation accordingly
    if openExcel is None:
        print("Workbook 'Hazmat_tracker' is not open.")
    else:
        print("Workbook 'Hazmat_tracker' found and connected.")
    # Specify the name of the sheet you want to work with
    sheet_name = "Spells"  # Change this to the name of your sheet

    # Get the sheet by name
    sheet = openExcel.sheets[sheet_name]
    return sheet, openExcel

def ReadSheet(sheet, openExcel):
    # Define your lookup value
    lookStow = 'Inbound Stow'
    lookDecant = 'Vendor Receive'
    lookSingles = 'Pack Singles'
    lookChute = 'Pack - Flow'
    lookSort = 'Sort - Flow'
    lookDock = 'Dock'
    lookPick = 'Pick'
    lookICQA = 'IC/QA/CS'
    lookFlexRT = 'FLEXRT'
    lookFLEXPT = 'FLEXPT'
    
    # Specifythe name of the sheet where you want to paste the information
    paste_sheet_name_stow = 'Inbound_Filtered'  # Change this to the name of your destination sheet
    paste_sheet_name_pick = 'Cap_Filtered'
    paste_sheet_name_ob = 'Outbound_Filtered'
    paste_sheet_name_flex = 'Flex'

    # Get the destination sheet by name
    paste_sheet_stow = openExcel.sheets[paste_sheet_name_stow]
    paste_sheet_cap = openExcel.sheets[paste_sheet_name_pick]
    paste_sheet_ob = openExcel.sheets[paste_sheet_name_ob]
    paste_sheet_flex = openExcel.sheets[paste_sheet_name_flex]

    found_IB = []
    found_CAP = []
    found_OB = []
    found_flex = []

    column_data = sheet.range((3, 17), (sheet.range('Q1').end('down').row, 18)) 
    
    for cell in column_data:
        #Inbound
        if (cell.value == lookStow or cell.value == lookDecant) and (sheet.cells(cell.row, 10).value == 'In'):
            found_row_data = [
            sheet.cells(cell.row, 7).value,  # Column 1
            sheet.cells(cell.row, 8).value,  # Column 2
            sheet.cells(cell.row, 13).value,  # Column 7
            sheet.cells(cell.row, 15).value,  # Column 8
            sheet.cells(cell.row, 17).value,  # Column 9
            sheet.cells(cell.row, 18).value  # Column 10
            ]
            found_IB.append(found_row_data)

            # Find the first empty row in the destination sheet
            next_row = paste_sheet_stow.cells(paste_sheet_stow.cells.last_cell.row, 1).end('up').row + 1
            # Paste the entire row onto the destination sheet
            paste_sheet_stow.range((next_row, 1)).value = found_row_data

        #CAP
        elif (cell.value == lookPick or cell.value == lookICQA) and (sheet.cells(cell.row, 10).value == 'In'):
            found_row_data = [
            sheet.cells(cell.row, 7).value,  # Column 1
            sheet.cells(cell.row, 8).value,  # Column 2
            sheet.cells(cell.row, 13).value,  # Column 7
            sheet.cells(cell.row, 15).value,  # Column 8
            sheet.cells(cell.row, 17).value,  # Column 9
            sheet.cells(cell.row, 18).value  # Column 10
            ]
            found_CAP.append(found_row_data)  

            # Find the first empty row in the destination sheet
            next_row = paste_sheet_cap.cells(paste_sheet_cap.cells.last_cell.row, 1).end('up').row + 1

            # Paste the entire row onto the destination sheet
            paste_sheet_cap.range((next_row, 1)).value = found_row_data

        #Outbound
        elif (cell.value == lookSingles or cell.value == lookChute or cell.value == lookSort or cell.value == lookDock) and (sheet.cells(cell.row, 10).value == 'In'):
            found_row_data = [
            sheet.cells(cell.row, 7).value,  # Column 1
            sheet.cells(cell.row, 8).value,  # Column 2
            sheet.cells(cell.row, 13).value,  # Column 7
            sheet.cells(cell.row, 15).value,  # Column 8
            sheet.cells(cell.row, 17).value,  # Column 9
            sheet.cells(cell.row, 18).value  # Column 10
            ]
            found_OB.append(found_row_data)

            # Find the first empty row in the destination sheet
            next_row = paste_sheet_ob.cells(paste_sheet_ob.cells.last_cell.row, 1).end('up').row + 1

            # Paste the entire row onto the destination sheet
            paste_sheet_ob.range((next_row, 1)).value = found_row_data

        #Flex
        elif (cell.value == lookFlexRT or cell.value == lookFLEXPT) and (sheet.cells(cell.row, 10).value == 'In'):
            found_row_data = [
            sheet.cells(cell.row, 7).value,  # Column 1
            sheet.cells(cell.row, 8).value,  # Column 2
            sheet.cells(cell.row, 13).value,  # Column 7
            sheet.cells(cell.row, 15).value,  # Column 8
            sheet.cells(cell.row, 17).value,  # Column 9
            sheet.cells(cell.row, 18).value  # Column 10
            ]
            found_flex.append(found_row_data)

            # Find the first empty row in the destination sheet
            next_row = paste_sheet_flex.cells(paste_sheet_flex.cells.last_cell.row, 1).end('up').row + 1

            # Paste the entire row onto the destination sheet
            paste_sheet_flex.range((next_row, 1)).value = found_row_data

        else:
            if (sheet.cells(cell.row, 10).value == 'OUT' or sheet.cells(cell.row, 10).value == 'zzz...'):
                break

    return found_IB, found_CAP, found_OB, found_flex, openExcel, sheet

def send_payload(ib_tabulated,ob_tabulated, CAP_tabulated,flexRT_tabulated):
    # Replace 'webhook_url' with the actual URL of your webhook
    # Example payload
    payload_format = {
        "Flex": flexRT_tabulated,  # Assign the tabulated data to the 'text' field
        "Inbound" : ib_tabulated,
        "Pick" : CAP_tabulated,
        "Outbound" : ob_tabulated   
    }
    webhook_url = 'https://hooks.slack.com/workflows/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'

    # Send POST request with JSON payload
    response = requests.post(webhook_url, json=payload_format)

    # Check if the request was successful (status code 200)
    if response.status_code == 200:
        print("Payload sent successfully.")
    else:
        print(f"Failed to send payload. Status code: {response.status_code}")

def main():
    sheet, openExcel = findOpenBook()
    found_IB, found_CAP, found_OB, found_flex, openExcel, sheet = ReadSheet(sheet, openExcel)

    ib_tabulated = tabulate(found_IB,headers=["Login", "Name", "Location", "Compliance", "Dept", "Cohort"], tablefmt="rounded_outline")
    ob_tabulated = tabulate(found_OB,headers=["Login", "Name", "Location", "Compliance", "Dept", "Cohort"], tablefmt="rounded_outline")
    CAP_tabulated = tabulate(found_CAP,headers=["Login", "Name", "Location", "Compliance", "Dept", "Cohort"], tablefmt="rounded_outline")
    flex_tabulated = tabulate(found_flex,headers=["Login", "Name", "Location", "Compliance", "Dept", "Cohort"], tablefmt="rounded_outline")
    
    #send_payload(ib_tabulated,ob_tabulated, CAP_tabulated,flex_tabulated)


if __name__ == "__main__":
    start_time = time.time()
    main()
    end_time = time.time()
    print(f"Execution time for theChef_2.py: {end_time - start_time} seconds")