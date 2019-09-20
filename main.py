from src.plan import FiscalPlan
from src.tke import TKE

import openpyxl
import win32com.client

if __name__ == "__main__":
    print("I`m main file")

    #excelBook = FiscalPlan(r"FiscalPlan")
    # excelBook.ppp()
    #excelBook.run()

    tke = TKE(r"TKE")
    tke.run()

    # excelApp = win32com.client.Dispatch("Excel.Application")
    # excelApp.Visible = False
    # wb0 = excelApp.Workbooks.Open(r"C:\Users\LuzhanskyiM-Inet\Development\Excel\TKE\Новый отчет.xls")
    # wb0.SaveAs(r"C:\Users\LuzhanskyiM-Inet\Development\Excel\TKE\Новый отчет1.xlsx", FileFormat=51)

    # wb = openpyxl.load_workbook(r"C:\Users\LuzhanskyiM-Inet\Development\Excel\TKE\Новый отчет1.xlsx", data_only=False)
    # wb.save(r"C:\Users\LuzhanskyiM-Inet\Development\Excel\TKE\Новый отчет2.xlsx")

    # wb1 = openpyxl.load_workbook(r"C:\Users\LuzhanskyiM-Inet\Development\Excel\TKE\Новый отчет1.xlsx", data_only=True)
    # wb1.save(r"C:\Users\LuzhanskyiM-Inet\Development\Excel\TKE\Новый отчет3.xlsx")
    

    # wb = openpyxl.load_workbook(r"C:\Users\LuzhanskyiM-Inet\Development\Excel\111.xlsx")
    # ws = wb["1"]
    # print(ws.cell(row=1,column=1).value)
    # # ws.cell(row=1,column=1).value = 147
    # # print(ws.cell(row=1,column=1).value)
    # wb.close()
    # del wb
    # # ws.cell(row=1,column=1).value = 4444447
    # print(ws.cell(row=1,column=1).value)
    # ws = wb["1"]
    # print(ws.cell(row=1,column=2).value)
