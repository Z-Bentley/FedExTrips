# win32 is unuseable with outlook so this will add the sort time
# to the clipboard after calculating the given times

import openpyxl
from datetime import datetime
import CustomizeExcel as CE

# filePath = 'Excel-Documents\\Sort_Time.xlsx'
# wb = openpyxl.load_workbook(filePath)

# Time Subtraction
def subtractTimes(time1, time2):
    format = '%H:%M'
    
    t1 = datetime.strptime(time1, format)
    t2 = datetime.strptime(time2, format)

    difference = (t2 - t1).total_seconds() // 60

    inMin = f"{difference:+.0f}"

    return inMin

# Local Sort Plan
def calcSortTimes(sheet, schTimes, actualTimes):
    try:
        # Set Column Headings
        sheet['A1'] = 'Flight 1460'
        sheet['B1'] = 'Schedule'
        sheet['C1'] = 'Actual' 
        sheet['D1'] = 'Variance'

        # Set Row Headings
        sheet['A2'] = 'Aircraft Arrival'
        sheet['A3'] = 'Sort Time'
        sheet['A4'] = 'Sort End'

        # Set Scheduled Times
        cells = ['B2', 'B3', 'B4']
        for bCell, time in zip(cells, schTimes):
            sheet[bCell] = time
        
        # Set Actual Times
        cells = ['C2', 'C3', 'C4']
        for c, t in zip(cells, actualTimes):
            sheet[c] = t
        
        # Time Math
        variCells = ['D2', 'D3', 'D4']
        for sch, act, cell in zip(schTimes, actualTimes, variCells):
            vari = subtractTimes(sch, act)
            sheet[cell] = vari
        
        # Add border
        CE.addBorder(sheet, 'A1:D4')

        # Column size adjusted
        cols = ['A', 'B', 'C', 'D']
        for col in cols:
            CE.adjustCol(sheet, col)
    except Exception as e:
        print(f"Error in calcSortTimes: {e}")

def setRootCauseDelay(sheet, actuals):
    try:
        # ???
        sheet['F1'] = 'X'
        sheet['G1'] = 'Late aircraft'
        sheet['H1'] = 'X'
        sheet['I1'] = 'Excess Minisort'

        sheet['I3'] = 'Plan = 6650lbs'
        sheet['I4'] = f'Actual = {actuals[0]}'
        sheet['I5'] = 'Plan = 655 pieces'
        sheet['I6'] = f'Actual = {actuals[1]}'

        # Add Border
        CE.addBorder(sheet, 'F1:I6')

        # Column size adjusted
        cols = ['F', 'G', 'H', 'I']
        for col in cols:
            CE.adjustCol(sheet, col)
    except Exception as e:
        print(f"Error in setRootCauseDelay: {e}")

def outboundTruckRoutes(sheet, schTimes, actualTrucks):
    try:
        # Truck Routes
        kCells = ['K2', 'K3', 'K4', 'K5', 'K6', 'K7', 'K8', 'K9', 'K11', 'K12', 'K13']
        truckRoutes = ['OXD02', 'CVG10', 'CVG03', 'FFT02', 'CVG06', 'OXD04', 'LUK01', 'CVG02', 'Docs LUK77/CVG77/OXD77FFT77', 'CVG78 (DNCA)', 'FFT41 (PDJA)']
        for k, tr in zip(kCells, truckRoutes):
            sheet[k] = tr

        # Scheduled Times
        lCells = ['L2', 'L3', 'L4', 'L5', 'L6', 'L7', 'L8', 'L9', 'L11', 'L12', 'L13']
        sheet['L1'] = 'Schedule'
        for l, sch in zip(lCells, schTimes):
            sheet[l] = sch

        # Actual Times
        sheet['M1'] = 'Actual'
        mCells = ['M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M11', 'M12', 'M13']
        for m, at in zip(mCells, actualTrucks):
            sheet[m] = at

        # Variance Calcs
        sheet['N1'] = 'Variance'
        nCells = ['N2', 'N3', 'N4', 'N5', 'N6', 'N7', 'N8', 'N9', 'N11', 'N12', 'N13']
        for n, scht, tru in zip(nCells, schTimes, actualTrucks):
            vari = subtractTimes(scht, tru)
            sheet[n] = vari

        # Adding Border
        CE.addBorder(sheet, 'K1:N13')

        # Column size adjusted
        cols = ['K', 'L', 'M', 'N']
        for col in cols:
            CE.adjustCol(sheet, col)
    except Exception as e:
        print(f"Error in outboundTruckRoutes: {e}")

# def runExample():
#     sheet = "Excel-Documents\\Sort_Time.xlsx"

#     scheduledSort = ["06:02", "06:26", "06:46"]
#     actualArrival = ["06:46", "06:46", "06:46"]

#     truckRoutes = ['OXD02', 'CVG10', 'CVG03', 'FFT02', 'CVG06', 'OXD04', 'LUK01', 'CVG02', 'Docs LUK77/CVG77/OXD77FFT77', 'CVG78 (DNCA)', 'FFT41 (PDJA)']
#     schudeldTimes = ['06:35', '07:25', '06:45', '07:15', '06:55', '07:00', '07:10', '07:05', '06:30', '07:00', '07:20']
#     actualTimes = ['09:05', '09:25', '09:05', '07:22', '07:00', '07:22', '07:05', '07:25', '07:05', '07:20', '07:20']
    
#     calcSortTimes(sheet, scheduledSort, actualArrival)
#     setRootCauseDelay(sheet, ['10856', '924 116 of this was NCING'])
#     outboundTruckRoutes(sheet, schudeldTimes, actualTimes)
#     print('Run it')

# runExample()