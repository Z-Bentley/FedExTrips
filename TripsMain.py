# Library Imports
import openpyxl
import subprocess
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

# File Imports
import WeightCalculations
import CalcSortTime
import copy_excel

##### Main file for running the Trips Program #####

# Create Window
def submit_data():
    data = {
        "localSchTimes": [schArrival.get(), schSortStart.get(), schSortEnd.get()],
        "localActTimes": [actArrival.get(), actSortStart.get(), actSortEnd.get()],
        "outSchTimes": [var.get() for var in schTimesVars],
        "outActTimes": [var.get() for var in truckTimesVars]
    }
    return data

# Browse for Excel File
def browseFiles():
    global filePath, sortTimeSheet, wb, sheet_created

    # Track if the sheet has been created already
    sheet_created = False

    filePath = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if filePath:
        fileSelectorButton.config(text=f'Selected File: {filePath}')
        wb = openpyxl.load_workbook(filePath)
        sheetName = 'sort_times'

        # Delete 'sort_times' sheet if it exists
        if sheetName in wb.sheetnames:
            wb.remove(wb[sheetName])
            print(f"Existing sheet '{sheetName}' has been deleted.")

        # Create a new 'sort_times' sheet only once
        if not sheet_created:
            sortTimeSheet = wb.create_sheet(sheetName)
            sheet_created = True
            print(f"New sheet '{sheetName}' has been created.")
    else:
        fileSelectorButton.config(text='No file selected')
    
def subButton():
    global sortTimeSheet

    # Prevent creating the sheet again if already processed
    if 'sort_times' not in wb.sheetnames:
        sortTimeSheet = wb.create_sheet('sort_times')
        print("Created new 'sort_times' sheet.")
    else:
        print("'sort_times' sheet already exists. Skipping creation.")

    # Submit Data
    data = submit_data()
    try:
        # Perform operations
        CalcSortTime.calcSortTimes(filePath, sortTimeSheet, data['localSchTimes'], data['localActTimes'])
        CalcSortTime.setRootCauseDelay(filePath, sortTimeSheet, ['10856', '924 116 of this was NCING'])
        CalcSortTime.outboundTruckRoutes(filePath, sortTimeSheet, data['outSchTimes'], data['outActTimes'])

        # Save and reopen
        wb.save(filePath)
        subprocess.Popen(['start', 'excel', filePath], shell=True)

        # Call external script
        copy_excel.copyExcel(filePath)
    except Exception as e:
        print(f"An error occurred: {e}")

# Setup window
root = tk.Tk()
root.title("FedEx Trips Reformer")
root.geometry("600x600")

fileSelectorButton = ttk.Label(root, text='No file selected', foreground='blue')
fileSelectorButton.grid(row=1, column=0, columnspan=4, pady=5)
tk.Button(root, text="Browse", command=browseFiles).grid(row=0, column=2)

# Pre-filled data and fillable data for Aircraft arrival and Sort
schArrival = tk.StringVar(value="06:02")
schSortStart = tk.StringVar(value="06:26")
schSortEnd = tk.StringVar(value="06:46")
# Fillable by the user
actArrival = tk.StringVar(value="06:46")  
actSortStart = tk.StringVar(value="06:46")  
actSortEnd = tk.StringVar(value="06:46")  

# Input fields for Aircraft and Sort times
ttk.Label(root, text="Local Sort Plan").grid(row=2, column=0, pady=(20,5))

labels_texts = ["Aircraft Arrival", "Sort Start", "Sort End"]
variables = [(schArrival, actArrival), (schSortStart, actSortStart), (schSortEnd, actSortEnd)]

for i, text in enumerate(labels_texts):
    ttk.Label(root, text=text).grid(row=i+3, column=0)
    ttk.Entry(root, textvariable=variables[i][0]).grid(row=i+3, column=1)
    ttk.Entry(root, textvariable=variables[i][1]).grid(row=i+3, column=2)

# Truck routes data
truckRoutes = ['OXD02', 'CVG10', 'CVG03', 'FFT02', 'CVG06', 'OXD04', 'LUK01', 'CVG02', 'Docs LUK77/CVG77/OXD77FFT77', 'CVG78 (DNCA)', 'FFT41 (PDJA)']
schTimes = ['06:35', '07:25', '06:45', '07:15', '06:55', '07:00', '07:10', '07:05', '06:30', '07:00', '07:20']
# Example fillable
truckTimes = ['09:05', '09:25', '09:05', '07:22', '07:00', '07:22', '07:05', '07:25', '07:05', '07:20', '07:20']

# Variables to store scheduled and actual times
schTimesVars = [tk.StringVar(value=time) for time in schTimes]
truckTimesVars = [tk.StringVar(value=time) for time in truckTimes]

# Input fields for Truck routes
ttk.Label(root, text="Outbound Truck Routes").grid(row=5, column=0)

for i, route in enumerate(truckRoutes):
    rowNum = i + 6
    ttk.Label(root, text=route).grid(row=rowNum, column=0)
    ttk.Entry(root, textvariable=schTimesVars[i]).grid(row=rowNum, column=1)  # Scheduled times
    ttk.Entry(root, textvariable=truckTimesVars[i]).grid(row=rowNum, column=2)  # Actual times

# Submit button
ttk.Button(root, text="Submit", command=subButton).grid(row=rowNum + 1, column=1, pady=20)
root.mainloop()

# # Save the workbook (make sure the excel is closed)
wb.save(filePath)
print('Workbook has been saved!')