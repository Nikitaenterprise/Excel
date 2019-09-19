from src.plan import FiscalPlan
from src.tke import TKE

import openpyxl

if __name__ == "__main__":
    print("I`m main file")

    #excelBook = FiscalPlan(r"FiscalPlan")
    # excelBook.ppp()
    #excelBook.run()

    tke = TKE(r"TKE")
    tke.run()

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
