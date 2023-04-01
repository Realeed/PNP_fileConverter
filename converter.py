import csv

float_precision = 3

def getRow(sheet, cellName):
    for index, row in enumerate(sheet):
        if cellName in row:
            return index
        
def getColumn(sheet, cellName):
    for index, column in enumerate(sheet[getRow(sheet, cellName)]):
        if cellName in column:
            return index

def convertFile(path):
    fidX = []
    fidY = []
    fidQty = 0

    designators = []
    comments = []
    footprints = []
    compX = []
    compY = []
    rotations = []
    compQty = 0
    print(designators)
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

        for row in range(headerRow + 1, maxRow):
            if sheet[row][footprintCol] == 'Global_Fiducial' or sheet[row][footprintCol] == 'Global_Fiducial_-_SQ':
                fidX.append(round(float(sheet[row][footprintCol+1]), float_precision))
                fidY.append(round(float(sheet[row][footprintCol+2]), float_precision))

        fidQty = len(fidX)

        designators = []
        comments = []
        footprints = []
        compX = []
        compY = []
        rotations = []

        for row in range(headerRow + fidQty + 1, maxRow):
            designators.append(sheet[row][designatorCol])
            comments.append(sheet[row][commentCol])
            footprints.append(sheet[row][footprintCol])
            compX.append(round(float(sheet[row][xCol]), float_precision))
            compY.append(round(float(sheet[row][yCol]), float_precision))
            rotations.append(sheet[row][rotaionCol])

        compQty = len(compX)

    with open(f'{path}', 'w', newline='', encoding='utf-8') as out_file:
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
        out_file.writerow(['PCB', '', '', '', '', '', '', '', '', '0', '0', '', '1', '0', 
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
        for i in range(fidQty):
            out_file.writerow(['Mark', fidX[i], fidY[i], '0.8', '1.2', 'Inner', '20', '4', '80',
                        'Black Spot', 'Left Camera'])
        out_file.writerow([])

        out_file.writerow(['#Comp', 'Feeder ID', 'Comment', 'Footprint', 'Designator',
                        'Nozzle', 'Pos X', 'Pos Y', 'Angle', 'Skip', 'Position'])
        for i in range(compQty):
            out_file.writerow(['Comp', '', comments[i], footprints[i], designators[i], '', 
                            compX[i], compY[i], rotations[i], 'NO', 'Align'])