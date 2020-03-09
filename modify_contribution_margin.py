#!/usr/bin/env python

from excel import ExcelDocument
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

path = r'K:\Links\2020\Options\Contribution Margin Testing'
excel = ExcelDocument()



for filename in os.listdir(path):
    
        excel.open(path + "\\" + filename, 3) # open file and update links
        excel.set_visible(False)
        excel.display_alerts(False)
	
        excel.set_value("C9", "=IF(MIN(M5:CO5)<ABS(MAX(M5:CO5)),MIN(M5:CO5),MIN(M5:CO5))")
	
        excel.save()

excel.quit()

