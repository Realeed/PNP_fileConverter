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
        fileName = Path(path).stem.replace('Pick Place for ', '').replace('_N10', '').replace('Panel', 'Single')
        return fileName
    else:
        fileName = Path(path).stem.replace('Pick Place for ', '').replace('_N10', '')
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

def writeOutFile(outFilePath, rows, cols, fidQty, fidX, fidY, compQty, comments, footprints, designators,
                compX, compY, rotations):
    with open(outFilePath, 'w', newline='', encoding='utf-8') as out_file:
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
        out_file.writerow(['PCB', cols, rows, 325, 325, 325, 325, 325, 325, 0, 0, 1, 1, 0, 
                        0, '', 'NO', 1, '', '', '', 4, '', '', 325, 325, ''])
        out_file.writerow([])

        out_file.writerow(['#Panel', 'Pos X', 'Pos Y', 'Offset X', 'Offset Y', 'Angle',
                        'Skip', 'Position'])
        out_file.writerow([])

        out_file.writerow(['#Nozzle', 'NozzleID', 'Nozzle Type', 'Disabled'])
        out_file.writerow(['Nozzle', 1, '', 'NO'])
        out_file.writerow(['Nozzle', 2, '', 'NO'])
        out_file.writerow(['Nozzle', 3, '', 'NO'])
        out_file.writerow(['Nozzle', 4, '', 'NO'])
        out_file.writerow(['Nozzle', 5, '', 'NO'])
        out_file.writerow(['Nozzle', 6, '', 'NO'])
        out_file.writerow(['Nozzle', 7, '', 'NO'])
        out_file.writerow(['Nozzle', 8, '', 'NO'])
        out_file.writerow([])

        out_file.writerow(['#Mark', 'Pos X', 'Pos Y', 'Min Size', 'Max Size', 'Flash',
                        'Brightness', 'Searching Area', 'Circular Similarity',
                        'Nested Mode', 'Select Camera', 'Position'])
        for i in range(fidQty):
            out_file.writerow(['Mark', fidX[i], fidY[i], 0.8, 1.2, 'Inner', 20, 4, 80,
                        'Black Spot', 'Left Camera'])
        out_file.writerow([])

        out_file.writerow(['#Comp', 'Feeder ID', 'Comment', 'Footprint', 'Designator',
                        'Nozzle', 'Pos X', 'Pos Y', 'Angle', 'Skip', 'Position'])
        for i in range(compQty):
            out_file.writerow(['Comp', '', comments[i], footprints[i], designators[i], '', 
                            compX[i], compY[i], rotations[i], 'NO', 'Align'])
            
def correctCompOrder(designators, comments, footprints, compX, compY, rotations):
    for i in range(len(compX)):
        for j in range(len(compX)):
            if compX[j] > compX[i]:
                numDes = designators[i]
                designators[i] = designators[j]
                designators[j] = numDes

                numCom = comments[i]
                comments[i] = comments[j]
                comments[j] = numCom

                numFtp = footprints[i]
                footprints[i] = footprints[j]
                footprints[j] = numFtp

                numX = compX[i]
                compX[i] = compX[j]
                compX[j] = numX

                numY = compY[i]
                compY[i] = compY[j]
                compY[j] = numY

                numRot = rotations[i]
                rotations[i] = rotations[j]
                rotations[j] = numRot


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

    rows = 0
    cols = 0

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
    
    topOutFilePath = dir + fileName + '_Top_N10' + fileExt

    writeOutFile(topOutFilePath, rows, cols, topFidQty, topFidX, topFidY, topSingleCompQty, topSingleComments, topSingleFootprints,
                topSingleDesignators, topSingleCompX, topSingleCompY, topSingleRotations)
        

    if bottomCompQty > 0:
        correctFidOrder(bottomFidX, bottomFidY)

        bottomSingleCompQty = getSingleCompQty(bottomDesignators)
        appendSingleFirstCompXY(bottomCompQty, bottomSingleCompQty, bottomCompX, bottomCompY,
                            bottomSingleFirstCompX, bottomSingleFirstCompY)
        
        bottomSingleFirstCompXMinIndices = []
        bottomSingleFirstCompYMinIndices = []

        appendMinIndices(bottomSingleFirstCompX, bottomSingleFirstCompY, bottomSingleFirstCompXMinIndices, bottomSingleFirstCompYMinIndices)       

        bottomSingleFirstLowestIndex = 0
        bottomSingleFirstLowestIndex = getSingleFirstLowestIndex(bottomSingleFirstCompXMinIndices, bottomSingleFirstCompYMinIndices)
        bottomPanelFirstLowestIndex = 0
        bottomPanelFirstLowestIndex = getPanelFirstLowestIndex(bottomCompX, bottomCompY, bottomSingleFirstCompX, bottomSingleFirstCompY, bottomSingleFirstLowestIndex)

        appendSingleData(bottomPanelFirstLowestIndex, bottomSingleCompQty, bottomSingleDesignators, bottomDesignators,
                        bottomSingleComments, bottomComments, bottomSingleFootprints, bottomFootprints,
                        bottomSingleCompX, bottomCompX, bottomSingleCompY, bottomCompY, bottomSingleRotations, bottomRotations)

        bottomOutFilePath = dir + fileName + '_Bottom_N10' + fileExt

        writeOutFile(bottomOutFilePath, rows, cols, bottomFidQty, bottomFidX, bottomFidY, bottomSingleCompQty, bottomSingleComments, bottomSingleFootprints,
                    bottomSingleDesignators, bottomSingleCompX, bottomSingleCompY, bottomSingleRotations)

        return 2
    
    return 1
    

def convertPanel10(path):
    dir = makeDir('resources')

    fileName = changeFileName(path, False)
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

    topFidQty, bottomFidQty, topCompQty, bottomCompQty = getExcelData(path, topFidX, 
                    topFidY, bottomFidX, bottomFidY, topDesignators, bottomDesignators, 
                    topComments, bottomComments,topFootprints, bottomFootprints, topCompX, topCompY, 
                    bottomCompX, bottomCompY, topRotations, bottomRotations)

    if checkGetDataFailed(topCompQty):
        return 0
    
    correctFidOrder(topFidX, topFidY)

    correctCompOrder(topDesignators, topComments, topFootprints, topCompX, topCompY, topRotations)
    
    topOutFilePath = dir + fileName + '_Top_N10' + fileExt

    writeOutFile(topOutFilePath, 1, 1, topFidQty, topFidX, topFidY, topCompQty, topComments, topFootprints,
                topDesignators, topCompX, topCompY, topRotations)
    
    if bottomCompQty > 0:
        correctFidOrder(bottomFidX, bottomFidY)

        correctCompOrder(bottomDesignators, bottomComments, bottomFootprints, bottomCompX, bottomCompY, bottomRotations)

        bottomOutFilePath = dir + fileName + '_Bottom_N10' + fileExt

        writeOutFile(bottomOutFilePath, 1, 1, bottomFidQty, bottomFidX, bottomFidY, bottomCompQty, bottomComments, bottomFootprints,
                    bottomDesignators, bottomCompX, bottomCompY, bottomRotations)
        
        return 2
    
    return 1    