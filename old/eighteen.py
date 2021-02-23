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

path = r'K:\Links\2021\Yamaha Rigging'
excel = ExcelDocument()



for filename in os.listdir(path):
    if filename.endswith(r'.xlsx'):

        excel.open(path + "\\" + filename, 3) # open file and update links
        excel.set_visible(False)
        excel.display_alerts(False)
    
        excel.set_value("D12", "18")
        excel.set_value("M1", "18")
        
        excel.save()
        excel.close()

excel.quit()

