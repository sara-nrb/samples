import win32com.client
import os
import glob

xlapp = win32com.client.DispatchEx("Excel.Application")
wb = xlapp.workbooks.open(r"C:\Temp\py_excel\test")

#os.listdir(r"C:\Temp\py_excel\test\" + list + ".xlxs")
xlapp.Visible = True
wb.RefreshAll()
xlapp.CalculateUntilAsyncQueriesDone()
wb.Save()
xlapp.Quit()