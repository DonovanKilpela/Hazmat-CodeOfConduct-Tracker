# This script will paste the logins, complaince status 
# home department, Home area and Shift pattern into a Excel macro workbook 

import pandas as pd
import xlwings as xw
from tkinter import filedialog as fd


def hazmat_list():
    file_name = fd.askopenfilename(title="Select file", filetypes=(("CSV Files", "*.csv"),))
    if not file_name:
        print("Dialog box closed, without selection")
        return None
    else:
        new_file = pd.read_csv(file_name, usecols=[1, 2, 3, 7, 9], names=['Compliance Status', 
                                                                             'Home Department', 
                                                                             'Home Area',  
                                                                             'Shift Pattern', 
                                                                             'Login'], header=None)
    return new_file

def change_order(new_file):
    new_file = new_file[['Login', 'Compliance Status', 'Home Department',
                         'Home Area', 'Shift Pattern']]
    return new_file

def write_to_sheet(data, workbook_path):
    wb = xw.Book(workbook_path)
    ws = wb.sheets['Spells']  # Change sheet name if necessary
    ws.range('G2').options(index=False, header=False).value = data['Login']
    ws.range('O2').options(index=False, header=False).value = data['Compliance Status']
    ws.range('P2').options(index=False, header=False).value = data['Home Department']
    ws.range('Q2').options(index=False, header=False).value = data['Home Area']
    ws.range('R2').options(index=False, header=False).value = data['Shift Pattern']
    wb.save()


data = hazmat_list()
if data is not None:
    data = change_order(data)

    workbook_path = "Hazmat_tracker.xlsm"
    if workbook_path:
        write_to_sheet(data, workbook_path)


