# Library Imports
import openpyxl
import subprocess

# File Imports
import WeightCalculations

# Open outlook
# olPath = r"C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE"
# subprocess.Popen(olPath)


##### Main file for running the Trips Program #####
wb = openpyxl.load_workbook('Excel-Documents\\WBManifestTable_1706103354202.xlsx')
sheet = wb['FedEx Air Ops Workbench Report']

# Caluculate the Weight based off of Destination
def calcWeight(sheet):
    print("For which destination do you seek?")
    print(WeightCalculations.getDestOptions(sheet))
    dest = input('(not case sensitive)>>> ')
    print()
    upperDest = dest.upper()

    WeightCalculations.calcWeight(upperDest, sheet)

# calcWeight(sheet)