# win32 is unuseable with outlook so this will add the sort time
# to the clipboard after calculating the given times

import openpyxl
import WeightCalculations

wb = openpyxl.load_workbook('Excel-Documents\\WBManifestTable_1706103354202.xlsx')
sheet = wb['FedEx Air Ops Workbench Report']

tuple(sheet['B5':'E46'])

