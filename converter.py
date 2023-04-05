import csv
from pathlib import Path
import os

float_precision = 3

def getRow(sheet, cellName):
    for index, row in enumerate(sheet):
        if cellName in row:
            return index
        
def getColumn(sheet, cellName):
    for index, column in enumerate(sheet[getRow(sheet, cellName)]):
        if cellName in column:
            return index

def makeDir(folderName):
    dir = f'{folderName}\\'
    if not os.path.exists(dir):
        os.mkdir(dir)
    return dir

def changeFileName(path, single):
    if single:
       fileName = Path(path).stem.replace('Pick Place for ', '').replace('Panel', 'Single')
       return fileName
    else:
        fileName = Path(path).stem.replace('Pick Place for ', '')
        return fileName
    
def getFileExtension(path):
    fileExt = Path(path).suffix
    return fileExt

def getExcelData(path, topFidX, topFidY, bottomFidX, bottomFidY,
                topDesignators, bottomDesignators, topComments, bottomComments,
                topFootprints, bottomFootprints, topCompX, topCompY, 
                bottomCompX, bottomCompY, topRotations, bottomRotations):
    
    with open(path, 'r') as csv_file:
        sheet = list(csv.reader(csv_file))
        maxRow = len(sheet)
        headerRow = getRow(sheet, 'Designator')
        
        designatorCol = getColumn(sheet, 'Designator')
        commentCol = getColumn(sheet, 'Comment')
        footprintCol = getColumn(sheet, 'Footprint')
        xCol = getColumn(sheet, 'Center-X(mm)')
        yCol = getColumn(sheet, 'Center-Y(mm)')
        rotaionCol = getColumn(sheet, 'Rotation')
        layerCol = getColumn(sheet, 'Layer')

        for row in range(headerRow + 1, maxRow):
            if sheet[row][footprintCol] == 'Global_Fiducial' or sheet[row][footprintCol] == 'Global_Fiducial_-_SQ':
                if sheet[row][layerCol] == 'TopLayer':
                    topFidX.append(round(float(sheet[row][xCol]), float_precision))
                    topFidY.append(round(float(sheet[row][yCol]), float_precision))
                elif sheet[row][layerCol] == 'BottomLayer':
                    bottomFidX.append(-1 * round(float(sheet[row][xCol]), float_precision))
                    bottomFidY.append(round(float(sheet[row][yCol]), float_precision))

        topFidQty = len(topFidX)
        bottomFidQty = len(bottomFidX)

        for row in range(headerRow + 1, maxRow):
            if sheet[row][footprintCol] != 'Global_Fiducial' and sheet[row][footprintCol] != 'Global_Fiducial_-_SQ':
                if sheet[row][layerCol] == 'TopLayer':
                    topDesignators.append(sheet[row][designatorCol])
                    topComments.append(sheet[row][commentCol])
                    topFootprints.append(sheet[row][footprintCol])
                    topCompX.append(round(float(sheet[row][xCol]), float_precision))
                    topCompY.append(round(float(sheet[row][yCol]), float_precision))
                    topRotations.append(sheet[row][rotaionCol])
                elif sheet[row][layerCol] == 'BottomLayer':
                    bottomDesignators.append(sheet[row][designatorCol])
                    bottomComments.append(sheet[row][commentCol])
                    bottomFootprints.append(sheet[row][footprintCol])
                    bottomCompX.append(-1 * round(float(sheet[row][xCol]), float_precision))
                    bottomCompY.append(round(float(sheet[row][yCol]), float_precision))
                    bottomRotations.append(180 - round(float(sheet[row][rotaionCol]), float_precision))

    topCompQty = len(topCompX)
    bottomCompQty = len(bottomCompX)

    return topFidQty, bottomFidQty, topCompQty, bottomCompQty

def checkGetDataFailed(topCompQty):
    if topCompQty == 0:
        return 1
    
def correctFidOrder(fidX, fidY):
    for i in range(len(fidX)):
        for j in range(len(fidX)):
            if fidX[j] < fidX[i]:
                numX = fidX[i]
                fidX[i] = fidX[j]
                fidX[j] = numX

                numY = fidY[i]
                fidY[i] = fidY[j]
                fidY[j] = numY

def getSingleCompQty(designators):
    for i in range (1, len(designators)):
        if designators[i] == designators[0]:
            singleCompQty = i
            return singleCompQty

def appendSingleFirstCompXY(compQty, singleCompQty, compX, compY, singleFirstCompX, singleFirstCompY):
        for i in range (0, compQty, singleCompQty):
            singleFirstCompX.append(compX[i])
            singleFirstCompY.append(compY[i])

def appendMinIndices(singleFirstCompX, singleFirstCompY, singleFirstCompXMinIndices, singleFirstCompYMinIndices):
    for i in range (len(singleFirstCompX)):
        if singleFirstCompX[i] == min(singleFirstCompX):
            singleFirstCompXMinIndices.append(i)

    for i in range (len(singleFirstCompY)):
        if singleFirstCompY[i] == min(singleFirstCompY):
            singleFirstCompYMinIndices.append(i)

def getSingleFirstLowestIndex(singleFirstCompXMinIndices, singleFirstCompYMinIndices):
    for i in range (len(singleFirstCompXMinIndices)):
        if singleFirstCompXMinIndices[i] in singleFirstCompYMinIndices:
            singleFirstLowestIndex = singleFirstCompXMinIndices[i]
            return singleFirstLowestIndex
        
def getPanelFirstLowestIndex(compX, compY, singleFirstCompX, singleFirstCompY, singleFirstLowestIndex):
    for i in range(len(compX)):
        if compX[i] == singleFirstCompX[singleFirstLowestIndex] and compY[i] == singleFirstCompY[singleFirstLowestIndex]:
            panelFirstLowestIndex = i
            return panelFirstLowestIndex
        
def appendSingleData(panelFirstLowestIndex, singleCompQty, singleDesignators, designators, singleComments, comments,
                singleFootprints, footprints, singleCompX, compX, singleCompY, compY, singleRotations, rotations):
    for i in range(panelFirstLowestIndex, panelFirstLowestIndex + singleCompQty):
        singleDesignators.append(designators[i])
        singleComments.append(comments[i])
        singleFootprints.append(footprints[i])
        singleCompX.append(compX[i])
        singleCompY.append(compY[i])
        singleRotations.append(rotations[i])

def convertSingle10(path):
    dir = makeDir('resources')

    fileName = changeFileName(path, True)
    fileExt = getFileExtension(path)

    topFidX = []
    topFidY = []
    bottomFidX = []
    bottomFidY = []
    topFidQty = 0
    bottomFidQty = 0

    topDesignators = []
    bottomDesignators = []
    topComments = []
    bottomComments = []
    topFootprints = []
    bottomFootprints = []
    topCompX = []
    topCompY = []
    bottomCompX = []
    bottomCompY = []
    topRotations = []
    bottomRotations = []
    topCompQty = 0
    bottomCompQty = 0

    topSingleDesignators = []
    bottomSingleDesignators = []
    topSingleComments = []
    bottomSingleComments = []
    topSingleFootprints = []
    bottomSingleFootprints = []
    topSingleCompX = []
    topSingleCompY = []
    bottomSingleCompX = []
    bottomSingleCompY = []
    topSingleRotations = []
    bottomSingleRotations = []
    topSingleCompQty = 0
    bottomSingleCompQty = 0

    rows = 0
    cols = 0

    topSingleFirstCompX = []
    topSingleFirstCompY = []

    bottomSingleFirstCompX = []
    bottomSingleFirstCompY = []

    topFidQty, bottomFidQty, topCompQty, bottomCompQty = getExcelData(path, topFidX, 
                    topFidY, bottomFidX, bottomFidY, topDesignators, bottomDesignators, 
                    topComments, bottomComments,topFootprints, bottomFootprints, topCompX, topCompY, 
                    bottomCompX, bottomCompY, topRotations, bottomRotations)

    if checkGetDataFailed(topCompQty):
        return 0
    
    correctFidOrder(topFidX, topFidY)

    topSingleCompQty = getSingleCompQty(topDesignators)

    appendSingleFirstCompXY(topCompQty, topSingleCompQty, topCompX, topCompY,
                            topSingleFirstCompX, topSingleFirstCompY)

    
    topSingleFirstCompXMinIndices = []
    topSingleFirstCompYMinIndices = []

    appendMinIndices(topSingleFirstCompX, topSingleFirstCompY, topSingleFirstCompXMinIndices, topSingleFirstCompYMinIndices)    

    rows = len(topSingleFirstCompXMinIndices)
    cols = len(topSingleFirstCompYMinIndices)

    topSingleFirstLowestIndex = 0

    topSingleFirstLowestIndex = getSingleFirstLowestIndex(topSingleFirstCompXMinIndices, topSingleFirstCompYMinIndices)

    topPanelFirstLowestIndex = 0

    topPanelFirstLowestIndex = getPanelFirstLowestIndex(topCompX, topCompY, topSingleFirstCompX, topSingleFirstCompY, topSingleFirstLowestIndex)

    appendSingleData(topPanelFirstLowestIndex, topSingleCompQty, topSingleDesignators, topDesignators,
                    topSingleComments, topComments, topSingleFootprints, topFootprints,
                    topSingleCompX, topCompX, topSingleCompY, topCompY, topSingleRotations, topRotations)
        

    if bottomCompQty > 0:
        for i in range(len(bottomFidX)):
            for j in range(len(bottomFidX)):
                if bottomFidX[j] < bottomFidX[i]:
                    numX = bottomFidX[i]
                    bottomFidX[i] = bottomFidX[j]
                    bottomFidX[j] = numX

                    numY = bottomFidY[i]
                    bottomFidY[i] = bottomFidY[j]
                    bottomFidY[j] = numY

        for i in range (1, len(bottomDesignators)):
            if bottomDesignators[i] == bottomDesignators[0]:
                bottomSingleCompQty = i
                break

        for i in range (0, bottomCompQty, bottomSingleCompQty):
            bottomSingleFirstCompX.append(bottomCompX[i])
            bottomSingleFirstCompY.append(bottomCompY[i])
        
        bottomSingleFirstCompXMinIndices = []
        bottomSingleFirstCompYMinIndices = []

        for i in range (len(bottomSingleFirstCompX)):
            if bottomSingleFirstCompX[i] == min(bottomSingleFirstCompX):
                bottomSingleFirstCompXMinIndices.append(i)

        for i in range (len(bottomSingleFirstCompY)):
            if bottomSingleFirstCompY[i] == min(bottomSingleFirstCompY):
                bottomSingleFirstCompYMinIndices.append(i)       


        bottomSingleFirstLowestIndex = 0

        for i in range (len(bottomSingleFirstCompXMinIndices)):
            if bottomSingleFirstCompXMinIndices[i] in bottomSingleFirstCompYMinIndices:
                bottomSingleFirstLowestIndex = bottomSingleFirstCompXMinIndices[i]
                break

        bottomPanelFirstLowestIndex = 0

        for i in range(len(bottomCompX)):
            if bottomCompX[i] == bottomSingleFirstCompX[bottomSingleFirstLowestIndex] and bottomCompY[i] == bottomSingleFirstCompY[bottomSingleFirstLowestIndex]:
                bottomPanelFirstLowestIndex = i
                break

        for i in range(bottomPanelFirstLowestIndex, bottomPanelFirstLowestIndex + bottomSingleCompQty):
            bottomSingleDesignators.append(bottomDesignators[i])
            bottomSingleComments.append(bottomComments[i])
            bottomSingleFootprints.append(bottomFootprints[i])
            bottomSingleCompX.append(bottomCompX[i])
            bottomSingleCompY.append(bottomCompY[i])
            bottomSingleRotations.append(bottomRotations[i])

    
    with open(dir + fileName + '_Top_N10' + fileExt, 'w', newline='', encoding='utf-8') as out_file:
        out_file = csv.writer(out_file)
        out_file.writerow(['#Feeder', 'Feeder ID', 'Skip', 'Pos X', 'Pos Y', 'Angle', 
                        'Footprint', 'Comment', 'Nozzle', 'Pick Height', 'Pick delay', 
                        'Move Speed', 'Place Height', 'Place delay', 'Place Speed', 
                        'Accuracy', 'Width', 'Length', 'Thickness', 'Size Analyze', 
                        'Tray X', 'Tray Y', 'Columns', 'Rows', 'Right Top X', 'Right Top Y', 
                        'Vision Model', 'Brightness', 'Vision Error', 'Vision Flash',
                        'Feeder Type', 'NoisyPoint', 'Try times', 'Feed wait time', 
                        'Find Out Rectangle'])
        out_file.writerow([])

        out_file.writerow(['#PCB', 'Columns', 'Rows', 'Left Bottom X', 'Left Bottom Y', 
                        'Left Top X', 'Left Top Y', 'Right Top X', 'Right Top Y', 
                        'Mirror Board Left Bottom X', 'Mirror Board Left Bottom Y',
                        'Panelize Board Angle', 'Mirror Board', 'Marked Panel X',
                        'Marked Panel Y', 'Marked Panel Value', 'Manual Program',
                        'Feed PCB', 'Panelized Mark Point', 'PCB Width', 'PCB Length',
                        'Safe height', 'Manual Mark', 'Test', 'Detect X', 'Detect Y',
                        'Long PCB Input'])
        out_file.writerow(['PCB', cols, rows, '', '', '', '', '', '', '0', '0', '', '1', '0', 
                        '0', '', 'NO', '1', '', '', '', '4', '', '', '', '', ''])
        out_file.writerow([])

        out_file.writerow(['#Panel', 'Pos X', 'Pos Y', 'Offset X', 'Offset Y', 'Angle',
                        'Skip', 'Position'])
        out_file.writerow([])

        out_file.writerow(['#Nozzle', 'NozzleID', 'Nozzle Type', 'Disabled'])
        out_file.writerow(['Nozzle', '1', '', 'NO'])
        out_file.writerow(['Nozzle', '2', '', 'NO'])
        out_file.writerow(['Nozzle', '3', '', 'NO'])
        out_file.writerow(['Nozzle', '4', '', 'NO'])
        out_file.writerow(['Nozzle', '5', '', 'NO'])
        out_file.writerow(['Nozzle', '6', '', 'NO'])
        out_file.writerow(['Nozzle', '7', '', 'NO'])
        out_file.writerow(['Nozzle', '8', '', 'NO'])
        out_file.writerow([])

        out_file.writerow(['#Mark', 'Pos X', 'Pos Y', 'Min Size', 'Max Size', 'Flash',
                        'Brightness', 'Searching Area', 'Circular Similarity',
                        'Nested Mode', 'Select Camera', 'Position'])
        for i in range(topFidQty):
            out_file.writerow(['Mark', topFidX[i], topFidY[i], '0.8', '1.2', 'Inner', '20', '4', '80',
                        'Black Spot', 'Left Camera'])
        out_file.writerow([])

        out_file.writerow(['#Comp', 'Feeder ID', 'Comment', 'Footprint', 'Designator',
                        'Nozzle', 'Pos X', 'Pos Y', 'Angle', 'Skip', 'Position'])
        for i in range(topSingleCompQty):
            out_file.writerow(['Comp', '', topSingleComments[i], topSingleFootprints[i], topSingleDesignators[i], '', 
                            topSingleCompX[i], topSingleCompY[i], topSingleRotations[i], 'NO', 'Align'])
    
    if bottomCompQty > 0:
        with open(dir + fileName + '_Bottom_N10' + fileExt, 'w', newline='', encoding='utf-8') as out_file:
            out_file = csv.writer(out_file)
            out_file.writerow(['#Feeder', 'Feeder ID', 'Skip', 'Pos X', 'Pos Y', 'Angle', 
                            'Footprint', 'Comment', 'Nozzle', 'Pick Height', 'Pick delay', 
                            'Move Speed', 'Place Height', 'Place delay', 'Place Speed', 
                            'Accuracy', 'Width', 'Length', 'Thickness', 'Size Analyze', 
                            'Tray X', 'Tray Y', 'Columns', 'Rows', 'Right Top X', 'Right Top Y', 
                            'Vision Model', 'Brightness', 'Vision Error', 'Vision Flash',
                            'Feeder Type', 'NoisyPoint', 'Try times', 'Feed wait time', 
                            'Find Out Rectangle'])
            out_file.writerow([])

            out_file.writerow(['#PCB', 'Columns', 'Rows', 'Left Bottom X', 'Left Bottom Y', 
                            'Left Top X', 'Left Top Y', 'Right Top X', 'Right Top Y', 
                            'Mirror Board Left Bottom X', 'Mirror Board Left Bottom Y',
                            'Panelize Board Angle', 'Mirror Board', 'Marked Panel X',
                            'Marked Panel Y', 'Marked Panel Value', 'Manual Program',
                            'Feed PCB', 'Panelized Mark Point', 'PCB Width', 'PCB Length',
                            'Safe height', 'Manual Mark', 'Test', 'Detect X', 'Detect Y',
                            'Long PCB Input'])
            out_file.writerow(['PCB', cols, rows, '', '', '', '', '', '', '0', '0', '', '1', '0', 
                            '0', '', 'NO', '1', '', '', '', '4', '', '', '', '', ''])
            out_file.writerow([])

            out_file.writerow(['#Panel', 'Pos X', 'Pos Y', 'Offset X', 'Offset Y', 'Angle',
                            'Skip', 'Position'])
            out_file.writerow([])

            out_file.writerow(['#Nozzle', 'NozzleID', 'Nozzle Type', 'Disabled'])
            out_file.writerow(['Nozzle', '1', '', 'NO'])
            out_file.writerow(['Nozzle', '2', '', 'NO'])
            out_file.writerow(['Nozzle', '3', '', 'NO'])
            out_file.writerow(['Nozzle', '4', '', 'NO'])
            out_file.writerow(['Nozzle', '5', '', 'NO'])
            out_file.writerow(['Nozzle', '6', '', 'NO'])
            out_file.writerow(['Nozzle', '7', '', 'NO'])
            out_file.writerow(['Nozzle', '8', '', 'NO'])
            out_file.writerow([])

            out_file.writerow(['#Mark', 'Pos X', 'Pos Y', 'Min Size', 'Max Size', 'Flash',
                            'Brightness', 'Searching Area', 'Circular Similarity',
                            'Nested Mode', 'Select Camera', 'Position'])
            for i in range(bottomFidQty):
                out_file.writerow(['Mark', bottomFidX[i], bottomFidY[i], '0.8', '1.2', 'Inner', '20', '4', '80',
                            'Black Spot', 'Left Camera'])
            out_file.writerow([])

            out_file.writerow(['#Comp', 'Feeder ID', 'Comment', 'Footprint', 'Designator',
                            'Nozzle', 'Pos X', 'Pos Y', 'Angle', 'Skip', 'Position'])
            for i in range(bottomSingleCompQty):
                out_file.writerow(['Comp', '', bottomSingleComments[i], bottomSingleFootprints[i], bottomSingleDesignators[i], '', 
                                bottomSingleCompX[i], bottomSingleCompY[i], bottomSingleRotations[i], 'NO', 'Align'])

    if topCompQty > 0 and bottomCompQty > 0:
        return 2
    elif topCompQty > 0 and bottomCompQty == 0:
        return 1
    

def convertPanel10(path):
    pass