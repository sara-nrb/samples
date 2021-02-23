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

path = r'K:\Links\2020\Yamaha Rigging'
excel = ExcelDocument()



for filename in os.listdir(path):
    if filename.endswith(r'.xlsx'):

        excel.open(path + "\\" + filename, 3) # open file and update links
        excel.set_visible(False)
        excel.display_alerts(False)
    
        excel.set_value("C9", "=IF(MIN(M5:CO5)<ABS(MAX(M5:CO5)),MIN(M5:CO5),MIN(M5:CO5))")
        excel.set_value("P5", "=(($C$8*$I$11)-P8-P9-P15-P16)/($C$8*$I$11)")
        excel.set_value("T5", "=(($C$8*$I$11)-T8-T9-T15-T16)/($C$8*$I$11)")
        excel.set_value("X5", "=(($C$8*$I$11)-X8-X9-X15-X16)/($C$8*$I$11)")
        excel.set_value("AB5", "=(($C$8*$I$11)-AB8-AB9-AB15-AB16)/($C$8*$I$11)") 
        excel.set_value("AF5", "=(($C$8*$I$11)-AF8-AF9-AF15-AF16)/($C$8*$I$11)")
        excel.set_value("AJ5", "=(($C$8*$I$11)-AJ8-AJ9-AJ15-AJ16)/($C$8*$I$11)")
        excel.set_value("AN5", "=(($C$8*$I$11)-AN8-AN9-AN15-AN16)/($C$8*$I$11)")
        excel.set_value("AR5", "=(($C$8*$I$11)-AR8-AR9-AR15-AR16)/($C$8*$I$11)")
        excel.set_value("AV5", "=(($C$8*$I$11)-AV8-AV9-AV15-AV16)/($C$8*$I$11)")
        excel.set_value("AZ5", "=(($C$8*$I$11)-AZ8-AZ9-AZ15-AZ16)/($C$8*$I$11)")
        excel.set_value("BD5", "=(($C$8*$I$11)-BD8-BD9-BD15-BD16)/($C$8*$I$11)")
        excel.set_value("BH5", "=(($C$8*$I$11)-BH8-BH9-BH15-BH16)/($C$8*$I$11)")
        excel.set_value("BL5", "=(($C$8*$I$11)-BL8-BL9-BL15-BL16)/($C$8*$I$11)")
        excel.set_value("BP5", "=(($C$8*$I$11)-BP8-BP9-BP15-BP16)/($C$8*$I$11)")
        excel.set_value("BT5", "=(($C$8*$I$11)-BT8-BT9-BT15-BT16)/($C$8*$I$11)")
        excel.set_value("BX5", "=(($C$8*$I$11)-BX8-BX9-BX15-BX16)/($C$8*$I$11)")
        excel.set_value("CB5", "=(($C$8*$I$11)-CB8-CB9-CB15-CB16)/($C$8*$I$11)")
        excel.set_value("CF5", "=(($C$8*$I$11)-CF8-CF9-CF15-CF16)/($C$8*$I$11)")
        excel.set_value("CJ5", "=(($C$8*$I$11)-CJ8-CJ9-CJ15-CJ16)/($C$8*$I$11)")
        excel.set_value("CN5", "=(($C$8*$I$11)-CN8-CN9-CN15-CN16)/($C$8*$I$11)")
        
        excel.save()
        excel.close()

excel.quit()

