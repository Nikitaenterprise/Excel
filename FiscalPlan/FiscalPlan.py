import sys
#sys.path.append("..")

import openpyxl
import win32com.client

from Modules.ExcelBook import ExcelBook
#from Excel.Modules.Header import Header
sys.path.clear()

class FiscalPlan:
    
    def __init__(self, wbName):
        self.wb = ExcelBook(wbName, data_only=True)
    
    def run(self):

        return



if __name__ == "__main__":
    print("I`m FiscalPlan.py file")
    #run tests
