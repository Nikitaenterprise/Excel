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
        self.PAT = ExcelBook(PATExcel, read=False, keep_vba=True)
        self.SBUT = ExcelBook(SBUTExcel, read=False, keep_vba=True)
        self.TEZ = ExcelBook(TEZExcel, read=False, worksheet=2)
    
    def run(self):
        self.todayCash.readExcelFile()
        print("1: " + str(self.populationAndReligion()))
        print("2: " + str(self.teploenergy()))
        print("3: " + str(self.directContractIndustryEE()))
        print("4: " + str(self.directContractIndustryPR()))
        return

    def deleteFiles(self):
        """Deletes all created files with .xlsx extension
        """
        self.closeFiles()
        fileNameWithPathWithoutExtensionTEZ = os.path.splitext(self.TEZ.fileNameWithPath)[0]
        fileNameWithPathWithoutExtensionPAT = os.path.splitext(self.PAT.fileNameWithPath)[0]
        fileNameWithPathWithoutExtensionSBUT = os.path.splitext(self.SBUT.fileNameWithPath)[0]
        fileNameWithPathWithoutExtensionTodayCash = os.path.splitext(self.todayCash.fileNameWithPath)[0]
        os.remove(fileNameWithPathWithoutExtensionTEZ + ".xlsx")
        os.remove(fileNameWithPathWithoutExtensionPAT + ".xlsx")
        os.remove(fileNameWithPathWithoutExtensionSBUT + ".xlsx")
        os.remove(fileNameWithPathWithoutExtensionTodayCash + ".xlsx")
        return

    def end(self):
        self.closeFiles()
        return

    def closeFiles(self):
        """Closes files without saving
        """
        self.PAT.close()
        self.SBUT.close()
        self.TEZ.close()
        self.todayCash.close()
        return
    
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
                                row=kyivEnergoNotEeRow).value
                if "ЕЕ" in contractName:
                    self.kyivEnergoEeContractCash = cashValue
                elif "ЕЕ" not in contractName:
                    kyivEnergoNotEeCash += cashValue
        except:
            print("Нет категории: Енергетичні підприємства м.Києва")
            kyivEnergoNotEeCash = 0

        return teploenergyCash + kyivEnergoNotEeCash

    def directContractIndustryEE(self):
        """Finds sum of cash from direct contract with
        industries (EE) and TEZ companies
        """
        try:
            self.TEZ.readExcelFile()
            self.TEZ.unmerge()
            # Column C contain names and categories
            industryEeRow = self.todayCash.findCellByStr("2.2. Промисловість за прямими договорами", "C").row
            # Column J contain cash
            industryEeColumnWithCash = openpyxl.utils.column_index_from_string(str("J"))
            # Column E contain names of contracts wich were concluded with companies
            industryEeColumnWithNameOfContracts = openpyxl.utils.column_index_from_string(str("E"))
            # Column C contain names and categories
            industryEeColumnWithNameOfCompanyOrCategory = openpyxl.utils.column_index_from_string(str("C"))

            industryEeCash = 0
            TEZCash = 0
            while True:
                industryEeRow += 1
                categoryOrCompanyName = self.todayCash.ws.cell(
                                column=industryEeColumnWithNameOfCompanyOrCategory, 
                                row=industryEeRow).value
                if categoryOrCompanyName == "Всього по теплоенергетиці":
                    break

                contractName = self.todayCash.ws.cell(
                                column=industryEeColumnWithNameOfContracts, 
                                row=industryEeRow).value
                if "ЕЕ" in contractName:
                    industryEeCash += self.todayCash.ws.cell(
                                column=industryEeColumnWithCash, 
                                row=industryEeRow).value                    
                else:
                    # Search the ТЕЦ.xlsx file for companies overlaping
                    # and if so, cash of this companies would be added
                    try:
                        # Get cell from ТЕЦ.xlsx with company name that is equal to 
                        # company name in НадходженняНаКР_.xlsx
                        cell = self.TEZ.findCellByStr(str(categoryOrCompanyName), "D")
                        if cell == None:
                            raise AttributeError(cell)

                        cellValue = cell.value
                    except AttributeError:
                        # If cell have None type then there is no such company in 
                        # ТЕЦ.xlsx and thats why cellValue should be equal to 0
                        cellValue = 0

                    if cellValue != 0:
                        TEZCash += self.todayCash.ws.cell(
                                column=industryEeColumnWithCash, 
                                row=industryEeRow).value
        except:
            print("Нет категории: Промисловість за прямими договорами (ЕЕ)")
            industryEeCash = 0
            TEZCash = 0

        try:
            # Check is this vatiable is created
            # in teploenergy()
            kyivEnergoEeContractCash
        except:
            print("Нет категории: Енергетичні підприємства м.Києва (ЕЕ)")
            kyivEnergoEeContractCash = 0

        return industryEeCash + TEZCash + kyivEnergoEeContractCash

    def directContractIndustryPR(self):
        """Finds cash from direct contract with
        industries (PR) without cash from TEZ companies and
        SBUT companies and PAT companies
        """
        try:
            self.PAT.readExcelFile()
            self.PAT.unmerge()
            self.SBUT.readExcelFile()
            self.SBUT.unmerge()

            # Column C contain names and categories
            industryPrRow = self.todayCash.findCellByStr("2.2. Промисловість за прямими договорами", "C").row
            # Column J contain cash
            industryPrColumn = openpyxl.utils.column_index_from_string(str("J"))
            # Column E contain names of contracts wich were concluded with companies
            industryPrColumnWithNameOfContracts = openpyxl.utils.column_index_from_string(str("E"))
            # Column C contain names and categories
            industryPrColumnWithNameOfCompanyOrCategory = openpyxl.utils.column_index_from_string(str("C"))

            industryPrCash = 0
            while True:
                industryPrRow += 1
                categoryOrCompanyName = self.todayCash.ws.cell(
                                column=industryPrColumnWithNameOfCompanyOrCategory, 
                                row=industryPrRow).value

                if categoryOrCompanyName == "Всього по теплоенергетиці":
                    break

                contractName = self.todayCash.ws.cell(
                                column=industryPrColumnWithNameOfContracts, 
                                row=industryPrRow).value
                if "ПР" in contractName:
                    # Search in:
                    # ТЕЦ.xlsx, 
                    # Всі Категорії. (без спожив. за )_ЗБУТ.xlsx, 
                    # Всі Категорії. (без спожив. за )_ПАТ.xlsx 
                    # files for companies overlaping and if so, 
                    # cash of this companies wouldn`t be calculated
                    try:
                        cellTEZ = self.TEZ.findCellByStr(str(categoryOrCompanyName), "D")
                        cellPAT = self.PAT.findCellByStr(str(categoryOrCompanyName), "C")
                        cellSBUT = self.SBUT.findCellByStr(str(categoryOrCompanyName), "C")
                        if cellTEZ != None or cellPAT != None or cellSBUT != None:
                            continue
                        # if cellTEZ == None or cellPAT == None or cellSBUT == None:
                        #     raise AttributeError()
                    except AttributeError:
                        continue
                        
                    industryPrCash += self.todayCash.ws.cell(
                            column=industryPrColumn, 
                            row=industryPrRow).value                    
        except:
            print("Нет категории: Промисловість за прямими договорами (ПР)")
            industryPrCash = 0


        return industryPrCash
            

if __name__ == "__main__":
    print("I`m FiscalPlan.py file")
    #run tests
