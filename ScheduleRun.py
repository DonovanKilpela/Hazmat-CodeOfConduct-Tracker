import schedule
import xlwings as xw
import os
import time
import subprocess

# Function to execute your module
def run_module():
    home_directory = os.path.expanduser("~")
    target_folder = 'Desktop/Hazmat_Tracker/dist/RunHazmat.exe'
    file_path = os.path.join(home_directory,target_folder)
    subprocess.run([file_path])

def clear_module():
    wb = xw.Book("Hazmat_tracker.xlsm")
    macro1 = wb.macro("InOutTimeCalc.Reset")
    macro1()

# Schedule the task to run three times a day
schedule.every().day.at("08:00").do(run_module)  
schedule.every().day.at("12:00").do(run_module)
schedule.every().day.at("16:00").do(run_module)
schedule.every().day.at("16:30").do(clear_module)

#schedule.every(10).minutes.do(run_module)

# Main loop to keep the script running
while True:
    schedule.run_pending()
    time.sleep(10)  # Check every minute if there's a scheduled task to run