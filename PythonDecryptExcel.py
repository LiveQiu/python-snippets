from win32com.client import Dispatch
import os

xlapp = Dispatch("Excel.Application")
Src = "C:/Users/Desktop/1.xls"

for i in range(999999):
    pwd = str(i)
    try:
        wb = xlapp.Workbooks.Open(Src, False, True, None, pwd)
        xlapp.DisplayAlerts = True
        print("Right password![%s]" % pwd)
        break
    except Exception as e:
        print("wrong password![%s]" % pwd)

wb.Close()
xlapp.Quit()