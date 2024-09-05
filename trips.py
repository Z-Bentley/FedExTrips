###### Outline of Trips program
import openpyxl

wb = openpyxl.load_workbook('Excel-Documents\\WBManifestTable_1706103354202.xlsx')
sheet = wb['FedEx Air Ops Workbench Report']

tuple(sheet['B5':'E46'])


# Gets ULD number from the Excel sheet
def getCanNum(cans):
    arrayOfCans = []

    for can in cans:
        arrayOfCans.append(can.value)
    
    print(arrayOfCans)
    return arrayOfCans

# Cut the can number to the type of can which is always the first 3 characters
def checkUldType(arrayOfCans):
    typesOfCans = []
    
    for canNum in arrayOfCans:
        typesOfCans.append(canNum[:3])

    return typesOfCans

# Weights of each can
def weightOfCan(type):
    canWeight = 0
    if type == 'AAD':
        canWeight = 500
    elif type == 'AMJ':
        canWeight = 400
    elif type == 'TRK':
        canWeight = 800
    elif type == 'AYY':
        canWeight = 200
    elif type == 'PMC':
        canWeight = 4
    elif type == 'AKE':
        canWeight = 50
    elif type == 'AAX':
        canWeight = 600

    return canWeight

def getWeight(weightOfCans):
    sumTotal = 0

    for cell in weightOfCans:
        if cell.value:
            sumTotal += int(cell.value)
    
    return sumTotal

def getDestCans(dest):
    destCoords = []

    for rowOfCellObjects in sheet['E5':'E46']:
        for cellObj in rowOfCellObjects:
            if cellObj.value == dest:
                rowNum = cellObj.coordinate[1:]

                destCoords.append(rowNum)

    return destCoords

# Main function
def calcWeight(dest):
    destCans = getDestCans(dest)

    weightCells = [sheet[f'D{num}'] for num in destCans]

    sumTotal = getWeight(weightCells)
    print(f'With can weight, the total weight of freight is {sumTotal}')

    canNumCells = [sheet[f'B{num}'] for num in destCans]
    cantypes = getCanNum(canNumCells)

    ulds = checkUldType(cantypes)

    for uld in ulds:
        sumTotal -= weightOfCan(uld)
    
    print(f'After complicated math here is the total freight weight without those pesky cans: {sumTotal}')
    return sumTotal

# for a specific uld destination
dest = input("For which destination do you seek?\n('FFTA', 'CVGA', 'LUKA')>>> ")
# dest = 'cvgrt'
upperDest = dest.upper()

calcWeight(upperDest)