import win32com.client as win32com
import time

Files = [r'C:\Temp\py_excel\test\1760.xlsx', 'C:\Temp\py_excel\test\1761.xlsx', 'C:\Temp\py_excel\test\1763.xlsx', 'C:\Temp\py_excel\test\1800.xlsx', 'C:\Temp\py_excel\test\1801.xlsx', 'C:\Temp\py_excel\test\1805.xlsx', 'C:\Temp\py_excel\test\1810 - CO.xlsx']

def RefreshFiles(Files):
	UpdatedFiles = []
	
	xl = win32.DispatchEx("Excel.Application")
	xl.visible = True
	for filename in Files:
		try:
			WB = xl.workbooks.open(filename)
			for con in WB.Connections:
				RefreshAllData
				time.sleep(5)
				WB.Close(True)
				
			UpdatedFiles.append(filename)
			time.sleep(1)
		except:
			print ("File Not Updated : " + filename)
			
	xl.Quit()