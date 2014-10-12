#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Erik
#
# Created:     09-10-2014
# Copyright:   (c) Erik 2014
# Licence:     Drivers
#-------------------------------------------------------------------------------

import xlrd
import numpy as np
from operator import itemgetter
import string

#set up sheet to take data from with xlrd
file = "Projections.xlsx"
workbook = xlrd.open_workbook(file)
sheet = workbook.sheet_by_index(0)
#gets number or rows from sheet to be used for size alocation
size = sheet.nrows
#creates 2d array with a string and a float with numpy (rankings = [(name,fpts),(name2,fpts)....])
rankings = np.zeros((size),dtype=('a16, f8'))  

# Set these values to the values for your league
fgVal = 1  #assuming made and missed field goals have the same values this can stay as 1
ftmadeVal = 1
ftmissVal = 1.5
tpmVal = 1 #three point made
rebVal = 1
astVal = 1
stlVal = 1.5
blkVal = 1.5
ptsVal = 1




for row in range(size - 1): #assumes the excel file has a first line of labels (player rank, name, fg, etc)
    row = row +1;
    name = sheet.cell_value(row,1)
    name = name.encode("utf-8")
    fg = (sheet.cell_value(row,2) - (1 - sheet.cell_value(row,2)))*10
    ft = (sheet.cell_value(row,3) * ftmadeVal) - (1 - float(sheet.cell_value(row,3)) * ftmissVal)
    tpm = sheet.cell_value(row,4) * tpmVal
    reb = sheet.cell_value(row,5) * rebVal
    ast = sheet.cell_value(row,6) * astVal
    stl = float(sheet.cell_value(row,7)) * stlVal
    blk = float(sheet.cell_value(row,8)) * blkVal
    pts = sheet.cell_value(row,9) * ptsVal
    fpts = fg + ft + ft + tpm + reb + ast + stl + blk + pts
    name = string.split(name,",")[0]
    name = string.split(name, "*")[0]
    rankings[row-1][0] = name
    rankings[row-1][1] = round(fpts, 2)

rankings = sorted(rankings, key=itemgetter(1), reverse=True)
for row in range(size):
    print rankings[row]

out = open('rankings.txt', 'w+')

for rows in range(size):
    out.write(str(rows+1) + " " + str(rankings[rows]) + "\n")  
out.close()