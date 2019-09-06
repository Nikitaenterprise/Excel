import sys
sys.path.append("..")

import openpyxl
import win32com.client

from Modules.ExcelBook import ExcelBook
from Modules.Header import Header


class FiscalPlan:
    def __init__(self, wbName):
        self.wb = ExcelBook(wbName, data_only=True)
        


if __name__ == "__main__":
    print("I`m main file")
    first = FiscalPlan("C:/Users/LuzhanskyiM-Inet/Development/Fiscal plan/Всі Категорії. (без спожив. за СЕРПЕНЬ)_ЗБУТ.xlsx")
    print(first.wb.ws["C17"].value)
