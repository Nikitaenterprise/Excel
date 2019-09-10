import sys
#sys.path.append("..")

import openpyxl
import win32com.client

from Modules.ExcelBook import ExcelBook
#from Excel.Modules.Header import Header

class FiscalPlan:
    
    def __init__(self, fiscalPlanExcel: str, todayCashExcel: str, 
                PATExcel: str, SBUTExcel: str, TEZExcel: str):
        self.fiscalPlan = ExcelBook(fiscalPlanExcel, read=False)
        self.todayCash = ExcelBook(todayCashExcel, read=False)
        self.PAT = ExcelBook(PATExcel, read=False)
        self.SBUT = ExcelBook(SBUTExcel, read=False)
        self.TEZ = ExcelBook(TEZExcel, read=False)
    
    def run(self):
        self.todayCash.readExcelFile()
        print(self.PopulationAndReligion())


        return

    def PopulationAndReligion(self):
        """Fins sum of cash from population and religion
        """
        populationRow = self.todayCash.findCellByStr("1.2. Населення", "C").row
        populationColumn = openpyxl.utils.column_index_from_string(str("J"))
        populationCash = self.todayCash.ws.cell(column=populationColumn, row=populationRow).value

        religionRow = self.todayCash.findCellByStr("Релігійні організації", "C").row
        religionColumn = populationColumn # the same column
        religionCash = self.todayCash.ws.cell(column=religionColumn, row=religionRow).value

        return populationCash + religionCash

    def Teploenergy(self):
        """
        """
        teploenergyRow = self.todayCash.findCellByStr("3.2. Теплоенергетика за прямими договорами", "C").row
        teploenergyColumn = openpyxl.utils.column_index_from_string(str("J"))
        teploenergyCash = self.todayCash.ws.cell(column=teploenergyColumn, row=teploenergyRow).value

        kyivEnergoNotEeRow = 
        



if __name__ == "__main__":
    print("I`m FiscalPlan.py file")
    #run tests
