###### Outline of Trips program
import openpyxl

wb = openpyxl.load_workbook('Excel-Documents\\WBManifestTable_1706103354202.xlsx')
sheet = wb['FedEx Air Ops Workbench Report']


# Gets ULD number from the Excel sheet
def getCanNum():
    arrayOfCans = []

    for rowOfCellObjects in sheet['B5':'B46']:
        for cellObj in rowOfCellObjects:
            if cellObj.value == 'None':
                break
            else:
                arrayOfCans.append(cellObj.value)
    
    return arrayOfCans

# Cut the can number to the type of can which is always the first 3 characters
def checkUldType(arrayOfCans):
    typesOfCans = []
    
    for canNum in arrayOfCans:
        type = canNum[:3]
        typesOfCans.append(type)

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

# Main function
def calcWeight():
    sumTotal = 0
    for rowOfCellObjects in sheet['D5':'D46']:
        for num in rowOfCellObjects:
            if num.value:
                weight = int(num.value)
                sumTotal += weight
    
    print('With can weight, the total weight of freight is %s' % sumTotal)

    cantypes = getCanNum()
    ulds = checkUldType(cantypes)
    for uld in ulds:
        sumTotal = sumTotal - weightOfCan(uld)
    
    print('After complicated math here is the total freight weight without those pesky cans: %s' % sumTotal)
    return sumTotal



calcWeight()