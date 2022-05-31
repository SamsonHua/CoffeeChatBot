#==============================================#
#                ATTA FIKA BOT                 #
#==============================================#
#       DEVELOPED BY CONRAD B & SAMSON H       #
#==============================================#
#    YOU MUST LOGIN TO FIKA@ATTABOTICS.COM     #
#               FOR THIS TO WORK               #
#==============================================# 

import numpy
import openpyxl
from pathlib import Path
import datetime
import os



#============================================================#
# >>> READ EXCEL SHEET "ThisWeeksMatches.xlsx" stored in the folder called Matches
#============================================================#
year = ""
month = ""
day = ""

matches = os.listdir('Matches')

if "ThisWeeksMatches" in matches[len(matches)-1]:
    year = matches[len(matches)-1].split('_')[1]
    month = matches[len(matches)-1].split('_')[2]
    day = matches[len(matches)-1].split('_')[3]

pathName = 'Matches/ThisWeeksMatches_%s_%s_%s' % (year,month,day)
print(pathName)
ExcelSheet = Path(pathName)
Workbook = openpyxl.load_workbook(ExcelSheet)
Sheet = Workbook.active

#============================================================#
# >>> FUNCTION TO READ PAIRINGS FROM EXCEL
#============================================================#

def getPairings():
    #Create 1x3 array

    Pairings = numpy.empty((0, 3))

    for row in Sheet.iter_rows():
        cellRow = numpy.empty(shape=(0,0))
        for cell in row:
            #Error check for a random blank space in 3rd column
            if cell.value == None:
                cellRow = numpy.append(cellRow, "NULL")
            else:
                cellRow = numpy.append(cellRow, cell.value)
        if cellRow.size < 3:
            #Error check if it is only two squares
            cellRow = numpy.append(cellRow, "NULL")

        Pairings = numpy.vstack([Pairings, cellRow])
    
    return Pairings
