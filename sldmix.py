# check for sld or mix cans
import openpyxl
import WeightCalculations as wCalc

wb = openpyxl.load_workbook('Excel-Documents\\WBManifestTable_1706103354202.xlsx')
sheet = wb['FedEx Air Ops Workbench Report']

tuple(sheet['B5':'E46'])

def checkForSldOrMix():
    canCoords = []
    for rowOfCellObjects in sheet['H5':'H46']:
        for cellObj in rowOfCellObjects:
            cellString = cellObj.value
            destCord = 'E' + cellObj.coordinate[1:]
            dest = sheet[destCord]
            if cellString and dest.value != 'CVGUP':
                if 'SLD' not in cellString and 'MIX' not in cellString:
                    coord = cellObj.coordinate[1:]
                    canCoords.append(coord)
    
    # print(canCoords)
    return canCoords

def getUlds(coords):
    ulds = []

    for num in coords:
        newCoord = 'B' + num
        uld = sheet[newCoord].value
        #   print(uld)
        ulds.append(uld)

    # print(ulds)
    return ulds

def getDests(coords):
    dests = []

    for num in coords:
        newCoord = 'E' + num
        dest = sheet[newCoord].value
        dests.append(dest)

    # print(dests)
    return dests

def getWeights(coords):
    weights = []

    for num in coords:
        newCoord = 'D' + num
        w = sheet[newCoord].value
        w = int(w)

        loadCord = 'H' + num
        loadString = sheet[loadCord].value

        if 'MST' in loadString:
            canCord = 'B' + num
            can = sheet[canCord].value
            # print(f"Original Weight: {w}")
            w -= wCalc.weightOfCan(can[:3])
            # print(f"AdjustedWeight: {w}")
        
        weights.append(w)

    return weights

# Objectify the uld number and dest
def createDict():
    cords = checkForSldOrMix()
    dests = getDests(cords)
    ulds = getUlds(cords)
    weights = getWeights(cords)

    uldObjs = []

    for uld, dest, weight in zip(ulds, dests, weights):
        uldDict = {
            'uld': uld,
            'dest': dest,
            'weight': weight
        }

        print(uldDict)
        uldObjs.append(uldDict)

    # print(uldObj)
    return uldObjs

cords = checkForSldOrMix()
# getUlds(cords)
# getDests(cords)
# getWeights(cords)


# createDict()
