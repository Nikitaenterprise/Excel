import os

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
        print(self.populationAndReligion())
        print(self.teploenergy())
        print(self.directContractIndustry())

        return

    def deleteFiles(self):
        os.remove(self.TEZ.fileNameWithPath)

    def populationAndReligion(self):
        """Finds sum of cash from population and religion
        """
        try:
            # Column C contain names and categories
            populationRow = self.todayCash.findCellByStr("1.2. Населення", "C").row
            # Column J contain cash
            populationColumn = openpyxl.utils.column_index_from_string(str("J"))
            populationCash = self.todayCash.ws.cell(column=populationColumn, row=populationRow).value
        except:
            print("Нет категории: Населення")
            populationCash = 0

        try:
            # Column C contain names and categories
            religionRow = self.todayCash.findCellByStr("Релігійні організації", "C").row
            # Column J contain cash
            religionColumn = populationColumn # the same column
            religionCash = self.todayCash.ws.cell(column=religionColumn, row=religionRow).value
        except:
            print("Нет категории: Релігійні організації")
            religionCash = 0

        return populationCash + religionCash

    def teploenergy(self):
        """Finds sum of cash from PSO users (ТЕ, КП, БО, РО)
        """
        try:
            # Column C contain names and categories
            teploenergyRow = self.todayCash.findCellByStr("3.2. Теплоенергетика за прямими договорами", "C").row
            # Column J contain cash
            teploenergyColumn = openpyxl.utils.column_index_from_string(str("J"))
            teploenergyCash = self.todayCash.ws.cell(column=teploenergyColumn, row=teploenergyRow).value
        except:
            print("Нет категории: Теплоенергетика за прямими договорами")
            teploenergyCash = 0

        try:
            # Column C contain names and categories
            kyivEnergoNotEeRow = self.todayCash.findCellByStr("Енергетичні підприємства м.Києва", "C").row
            # Column J contain cash
            kyivEnergoNotEeColumn = teploenergyColumn # the same column
            kyivEnergoNotEeCash = 0
            
            # Column E contain names of contracts wich were concluded with companies 
            kyivEnergoNotEeColumnWithNameOfContract = openpyxl.utils.column_index_from_string(str("E"))
            # Here, the contracts wich not contain EE are calculated
            for i in range(2, 6):
                kyivEnergoNotEeRow += i
                contractName = self.todayCash.ws.cell(
                                column=kyivEnergoNotEeColumnWithNameOfContract, 
                                row=kyivEnergoNotEeRow).value
                cashValue = self.todayCash.ws.cell(
                                column=kyivEnergoNotEeColumn, 
                                row=kyivEnergoNotEeRow)
                if "ЕЕ" in contractName:
                    self.kyivEnergoEeContractCash = cashValue
                elif "ЕЕ" not in contractName:
                    kyivEnergoNotEeCash += cashValue
        except:
            print("Нет категории: Енергетичні підприємства м.Києва")
            kyivEnergoNotEeCash = 0

        return teploenergyCash + kyivEnergoNotEeCash

    def directContractIndustry(self):
        """
        """
        try:
            self.TEZ.readExcelFile()
            self.TEZ.unmerge()
            # Column C contain names and categories
            industryEeRow = self.todayCash.findCellByStr("2.2. Промисловість за прямими договорами", "C").row
            # Column J contain cash
            industryEeColumn = openpyxl.utils.column_index_from_string(str("J"))
            # Column E contain names of contracts wich were concluded with companies
            industryEeColumnWithNameOfContracts = openpyxl.utils.column_index_from_string(str("E"))

            industryEeCash = 0
            TEZCash = 0
            i = 1
            while True:
                industryEeRow += i
                categoryOrCompanyName = self.todayCash.ws.cell(
                                column=industryEeColumn, 
                                row=industryEeRow).value

                if categoryOrCompanyName == "Всього по теплоенергетиці":
                    break

                contractName = self.todayCash.ws.cell(
                                column=industryEeColumnWithNameOfContracts, 
                                row=industryEeRow).value
                if "ЕЕ" in contractName:
                    industryEeCash += self.todayCash.ws.cell(
                                column=industryEeColumn, 
                                row=industryEeRow)                    

                # Search the ТЕЦ.xlsx file for companies overlaping
                # and if so, cash of this companies would be added 
                cellValue = self.TEZ.findCellByStr(str(categoryOrCompanyName), "D")
                if cellValue != None:
                    TEZCash += self.todayCash.ws.cell(
                                column=industryEeColumn, 
                                row=industryEeRow)
                i += 1
        except:
            print("Нет категории: Промисловість за прямими договорами (ЕЕ)")
            industryEeCash = 0
            TEZCash = 0

        try:
            # Check is this vatiable is created
            kyivEnergoEeContractCash
        except:
            print("Нет категории: Енергетичні підприємства м.Києва (ЕЕ)")
            kyivEnergoEeContractCash = 0

        
        return industryEeCash + TEZCash + kyivEnergoEeContractCash

if __name__ == "__main__":
    print("I`m FiscalPlan.py file")
    #run tests
