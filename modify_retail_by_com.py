#!/usr/bin/env python

from openpyxl import load_workbook
from excel import ExcelDocument
from pathlib import Path
import os

"""
 __   __  ___       __        _______   _____  ___    __    _____  ___    _______   
|"  |/  \|  "|     /""\      /"      \ (\"   \|"  \  |" \  (\"   \|"  \  /" _   "|  
|'  /    \:  |    /    \    |:        ||.\\   \    | ||  | |.\\   \    |(: ( \___)  
|: /'        |   /' /\  \   |_____/   )|: \.   \\  | |:  | |: \.   \\  | \/ \       
 \//  /\'    |  //  __'  \   //      / |.  \    \. | |.  | |.  \    \. | //  \ ___  
 /   /  \\   | /   /  \\  \ |:  __   \ |    \    \ | /\  |\|    \    \ |(:   _(  _| 
|___/    \___|(___/    \___)|__|  \___) \___|\____\)(__\_|_)\___|\____\) \_______) 

   - This will modify the spreadseets in "K:\Links\2020\Options"
   - Make backups before testing the sheets or at the very least limit testing
   - limit testing with "break" at end of for loop so only executes once
   - by setting debud = True

"""

debug = False
# load workbook, no need to close it is closed upon reading
wb = load_workbook("templates\revised_pricing.xlsx") 
ws = wb.active # selcet the active worksheet
# A/col_1 = Option, B/col_2 = Old_Retial, C/col_3 = New_Retail
# will access with ws.cell(row = row, column = column).value

excel = ExcelDocument(False) # open Excel do not use existing copy

for row, cell in enumerate(sh["A"], start =1):
    if row == 1: # only needed if ROW 1 has titles
        continue
    file = os.path.join(r"K:\Links\2020\Options\", cell.value +".xlsx")
    excel.open(file, 3) # open file and update links
    excel.set_visible(False)
    excel.display_alerts(False)
	
    # get all values from input sheet	
    new_retail = ws.cell(row = row, column = 3).value

    # write all values to output sheet
    excel.set_value("B11", new_retail)
	
    excel.save()
    # for debug purposes and only running once
    if debug:
        print(ws.cell(row = row, column = 1).value, ws.cell(row = row, column = 2).value, ws.cell(row = row, column = 3).value)
        break

excel.quit()

