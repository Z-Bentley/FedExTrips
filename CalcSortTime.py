# win32 is unuseable with outlook so this will add the sort time
# to the clipboard after calculating the given times

import openpyxl
from datetime import datetime
import CustomizeExcel as CE

# filePath = 'Excel-Documents\\Sort_Time.xlsx'
# wb = openpyxl.load_workbook(filePath)

# Time Subtraction
def subtractTimes(time1, time2):
    time_format = '%H:%M'

    # Parse the input times
    t1 = datetime.strptime(time1, time_format)
    t2 = datetime.strptime(time2, time_format)

    # Calculate the difference in minutes
    difference = int((t2 - t1).total_seconds() / 60)

    # Format the result as a string with "+" or "-"
    return f"+{difference}" if difference >= 0 else f"{difference}"

# Local Sort Plan
def calcSortTimes(sheet, schTimes, actualTimes):
    try:
        # Set Column Headings
        sheet.Cells(1, 1).Value = 'Flight 1460'
        sheet.Cells(1, 2).Value = 'Schedule'
        sheet.Cells(1, 3).Value = 'Actual'
        sheet.Cells(1, 4).Value = 'Variance'

        # Set Row Headings
        sheet.Cells(2, 1).Value = 'Aircraft Arrival'
        sheet.Cells(3, 1).Value = 'Sort Time'
        sheet.Cells(4, 1).Value = 'Sort End'

        # Populate Data and Variance
        for i in range(len(schTimes)):
            row = i + 2
            sheet.Cells(row, 2).Value = schTimes[i]  # Schedule
            sheet.Cells(row, 3).Value = actualTimes[i]  # Actual
            
            # Variance as a string
            variance = subtractTimes(schTimes[i], actualTimes[i])
            sheet.Cells(row, 4).Value = f"'{variance}"

        # Apply borders to the range
        CE.addBorder(sheet, 'A1:D4')
        print("Local Sort Plan calculated and set.")
    except Exception as e:
        print(f"Error in calcSortTimes: {e}")

def setRootCauseDelay(sheet, actuals):
    try:
        # Set Root Cause Delay Data
        sheet.Cells(7, 1).Value = 'X'
        sheet.Cells(8, 1).Value = 'Late aircraft'
        sheet.Cells(9, 1).Value = 'X'
        sheet.Cells(10, 1).Value = 'Excess Minisort'

        sheet.Cells(9, 4).Value = "Plan = 6650lbs"
        sheet.Cells(10, 4).Value = f"Actual = {actuals[0]}"
        sheet.Cells(11, 4).Value = "Plan = 655 pieces"
        sheet.Cells(12, 4).Value = f"Actual = {actuals[1]}"

        # Apply borders to the range
        CE.addBorder(sheet, 'A7:D12')
        print("Root Cause Delay set.")
    except Exception as e:
        print(f"Error in setRootCauseDelay: {e}")

def outboundTruckRoutes(sheet, schTimes, actualTrucks):
    try:
        # Truck Routes and Data
        truckRoutes = ['OXD02', 'CVG10', 'CVG03', 'FFT02', 'CVG06', 'OXD04', 
                       'LUK01', 'CVG02', 'Docs LUK77/CVG77/OXD77FFT77', 
                       'CVG78 (DNCA)', 'FFT41 (PDJA)']

        # Headers
        sheet.Cells(15, 1).Value = "Truck Route"
        sheet.Cells(15, 2).Value = "Schedule"
        sheet.Cells(15, 3).Value = "Actual"
        sheet.Cells(15, 4).Value = "Variance"

        # Populate Data and Variance
        for i in range(len(truckRoutes)):
            row = i + 16
            sheet.Cells(row, 1).Value = truckRoutes[i]  # Truck Route
            sheet.Cells(row, 2).Value = schTimes[i]  # Schedule
            sheet.Cells(row, 3).Value = actualTrucks[i]  # Actual
            
            # Variance as a string
            variance = subtractTimes(schTimes[i], actualTrucks[i])
            sheet.Cells(row, 4).Value = f"'{variance}"  # Add single quote to enforce string

        # Apply borders to the range
        CE.addBorder(sheet, 'A15:D27')
        print("Outbound Truck Routes set.")
    except Exception as e:
        print(f"Error in outboundTruckRoutes: {e}")

# print(subtractTimes("06:30", "07:15"))  # Output: "+45"
# print(subtractTimes("07:30", "06:45"))  # Output: "-45"
# print(subtractTimes("06:30", "06:30"))  # Output: "+0"