# win32 is unuseable with outlook so this will add the sort time
# to the clipboard after calculating the given times

import openpyxl
from datetime import datetime

filePath = 'Excel-Documents\\WBManifestTable_1706103354202.xlsx'
wb = openpyxl.load_workbook(filePath)
if 'sort_times' not in wb.sheetnames:
    sheet = wb.create_sheet("sort_times")
else:
    sheet = wb['sort_times']

# Time Subtraction
def subtractTimes(time1, time2):
    format = '%H:%M'
    t1 = datetime.strptime(time1, format)
    t2 = datetime.strptime(time2, format)

    isNegative = False
    if t2 >= t1:
        difference = t2 - t1
    else:
        difference = t1 - t2
        isNegative = True
    
    inMin = (difference.seconds // 60) % 60

    if inMin >= 0 and not isNegative:
        inMin = str(inMin)
        inMin = '+' + inMin
    else:
        inMin = str(inMin)
        inMin = '-' + inMin
    
    return inMin

# Local Sort Plan
def calcSortTimes(sheet, actualTimes):
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
    schArr = '06:02'
    schStart = '06:26'
    schEnd = '06:46'
    sheet['B2'] = schArr
    sheet['B3'] = schStart
    sheet['B4'] = schEnd
    
    scheduledTimes = [schArr, schStart, schEnd]
    
    # Set Actual Times
    cells = ['C2', 'C3', 'C4']
    for c, t in zip(cells, actualTimes):
        sheet[c] = t
    
    # Time Math
    variCells = ['D2', 'D3', 'D4']
    for sch, act, cell in zip(scheduledTimes, actualTimes, variCells):
        vari = subtractTimes(sch, act)
        sheet[cell] = vari
        # variances.append(vari)

def outboundTruckRoutes(sheet, actualTrucks):
    # Truck Routes
    sheet['A12'] = 'OXD02'
    sheet['A13'] = 'CVG10'
    sheet['A14'] = 'CVG03'
    sheet['A15'] = 'FFT02'
    sheet['A16'] = 'CVG06'
    sheet['A17'] = 'OXD04'
    sheet['A18'] = 'LUK01'
    sheet['A19'] = 'CVG02'
    sheet['A21'] = 'Docs LUK77/CVG77/OXD77FFT77'
    sheet['A22'] = 'CVG78 (DNCA)'
    sheet['A23'] = 'FFT41 (PDJA)'

    schTimes = ['06:35', '07:25', '06:45', '07:15', '06:55', '07:00', '07:10', '07:05', '06:30', '07:00', '07:20']
    bCells = ['B12', 'B13', 'B14', 'B15', 'B16', 'B17', 'B18', 'B19', 'B21', 'B22', 'B23']

    sheet['B11'] = 'Schedule'
    for b, sch in zip(bCells, schTimes):
        sheet[b] = sch

    # Actual Times
    sheet['C11'] = 'Actual'
    cCells = ['C12', 'C13', 'C14', 'C15', 'C16', 'C17', 'C18', 'C19', 'C21', 'C22', 'C23']
    for c, at in zip(cCells, actualTrucks):
        sheet[c] = at

    # Variance Calcs
    sheet['D11'] = 'Variance'
    dCells = ['D12', 'D13', 'D14', 'D15', 'D16', 'D17', 'D18', 'D19', 'D21', 'D22', 'D23']
    for d, scht, tru in zip(dCells, schTimes, actualTrucks):
        vari = subtractTimes(scht, tru)
        sheet[d] = vari



# arrival = input('What is the Arrival time? (hh:mm)')
arrival = '06:29'
# sortStart = input('When did the sort start? (hh:mm)')
sortStart = '06:43'
# sortEnd = input('When did the sort end? (hh:mm)')
sortEnd = '07:10'
actualTimes = [arrival, sortStart, sortEnd]
calcSortTimes(sheet, actualTimes)


truckTimes = ['07:05', '07:25', '07:05', '07:22', '07:00', '07:22', '07:05', '07:25', '07:05', '07:20', '07:20']
outboundTruckRoutes(sheet, truckTimes)


# Save the workbook (make sure the excel is closed)
wb.save(filePath)
print('Workbook has been saved!')