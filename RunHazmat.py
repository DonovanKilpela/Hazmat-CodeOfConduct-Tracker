import xlwings as xw
import subprocess

wb = xw.Book("Hazmat_tracker.xlsm")


theChef = 'theChef.exe'
macro1 = wb.macro("ActiveRoster.DoIt")
macro1()

subprocess.run([theChef])

macro2 = wb.macro("InOutTimeCalc.Reset_Filtered")
macro2()
