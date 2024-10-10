# Library Imports
import openpyxl
import subprocess
import tkinter as tk
from tkinter import ttk

# File Imports
import WeightCalculations
import CalcSortTime

##### Main file for running the Trips Program #####
filePath = 'Excel-Documents\\WBManifestTable_1706103354202.xlsx'
wb = openpyxl.load_workbook(filePath)
if 'sort_times' not in wb.sheetnames:
    sortTimeSheet = wb.create_sheet("sort_times")
else:
    sortTimeSheet = wb['sort_times']

# Create Window
def submit_data():
    localSchTimes = []
    localActTimes = []
    outSchTimes = []
    outActTimes = []
    
    # print("Local Sort Plan")
    # print(f"Aircraft Arrival: Scheduled - {schArrival.get()}, Actual - {actArrival.get()}")
    # print(f"Sort Start Time: Scheduled - {schSortStart.get()}, Actual - {actSortStart.get()}")
    # print(f"Sort End Time: Scheduled - {schSortEnd.get()}, Actual - {actSortEnd.get()}")

    localSchTimes.append(schArrival.get())
    localActTimes.append(actArrival.get())
    
    localSchTimes.append(schSortStart.get())
    localActTimes.append(actSortStart.get())
    
    localSchTimes.append(schSortEnd.get())
    localActTimes.append(actSortEnd.get())
    
    # Print and store Outbound Truck Routes times
    # print("\nOutbound Truck Routes:")
    for i, route in enumerate(truckRoutes):
        scheduled = schTimesVars[i].get()
        actual = truckTimesVars[i].get()
        # print(f"{route}: Scheduled - {scheduled}, Actual - {actual}")
        
        # Store the truck route times in arrays
        outSchTimes.append(scheduled)
        outActTimes.append(actual)

    # Now the submitted_schTimes and submitted_actTimes arrays contain all the data
    # print("\nSubmitted Scheduled Times:", outSchTimes)
    # print("Submitted Actual Times:", outActTimes)

    return localSchTimes, localActTimes, outSchTimes, outActTimes

root = tk.Tk()
root.title("FedEx Trips Reformer")
root.geometry("600x600")

# Pre-filled data and fillable data for Aircraft arrival and Sort
schArrival = tk.StringVar(value="06:02")
schSortStart = tk.StringVar(value="06:26")
schSortEnd = tk.StringVar(value="06:46")
actArrival = tk.StringVar(value="06:46")  # Fillable by the user
actSortStart = tk.StringVar(value="06:46")  # Fillable by the user
actSortEnd = tk.StringVar(value="06:46")  # Fillable by the user

# Input fields for Aircraft and Sort times
ttk.Label(root, text="Local Sort Plan").grid(row=0, column=0)

ttk.Label(root, text="Aircraft Arrival").grid(row=1, column=0)
ttk.Entry(root, textvariable=schArrival).grid(row=1, column=1)
ttk.Entry(root, textvariable=actArrival).grid(row=1, column=2)

ttk.Label(root, text="Sort Time").grid(row=2, column=0)
ttk.Entry(root, textvariable=schSortStart).grid(row=2, column=1)
ttk.Entry(root, textvariable=actSortStart).grid(row=2, column=2)

ttk.Label(root, text="Sort End").grid(row=3, column=0)
ttk.Entry(root, textvariable=schSortEnd).grid(row=3, column=1)
ttk.Entry(root, textvariable=actSortEnd).grid(row=3, column=2)

# Truck routes data
truckRoutes = ['OXD02', 'CVG10', 'CVG03', 'FFT02', 'CVG06', 'OXD04', 'LUK01', 'CVG02', 'Docs LUK77/CVG77/OXD77FFT77', 'CVG78 (DNCA)', 'FFT41 (PDJA)']
schTimes = ['06:35', '07:25', '06:45', '07:15', '06:55', '07:00', '07:10', '07:05', '06:30', '07:00', '07:20']
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


# # Scheduled and Actual Times
# localSch, localAct, outSch, outAct = submit_data()

# Caluculate the Weight based off of Destination
def calcWeight(sheet):
    print("For which destination do you seek?")
    print(WeightCalculations.getDestOptions(sheet))
    dest = input('(not case sensitive)>>> ')
    print()
    upperDest = dest.upper()

    WeightCalculations.calcWeight(upperDest, sheet)

# print(localSch)

# calcWeight(sheet)


def subButton():
    # Scheduled and Actual Times
    localSch, localAct, outSch, outAct = submit_data()
    actuals = ['10856', '924 116 of this was NCING']

    CalcSortTime.calcSortTimes(filePath, sortTimeSheet, localSch, localAct)
    CalcSortTime.setRootCauseDelay(filePath, sortTimeSheet, actuals)
    CalcSortTime.outboundTruckRoutes(filePath, sortTimeSheet, outSch, outAct)
    wb.save(filePath)
    print('Workbook has been saved!')

# Submit button
ttk.Button(root, text="Submit", command=subButton).grid(row=rowNum + 1, column=1)
root.mainloop()

# # Save the workbook (make sure the excel is closed)
# wb.save(filePath)
# print('Workbook has been saved!')