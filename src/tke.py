import os
import datetime
from copy import copy


import openpyxl
import win32com.client


from src.manager import *


class TKE:

    def __init__(self, dir: str):
        self.mng = Manager(os.path.abspath(dir))
        self.numberOfFilesToStart = 4
        self.checkIfDirectoryIsReady(dir)

    def checkIfDirectoryIsReady(self, path: str):

        self.mng.addFilesInDir()

        self.mng.getFile("Новый отчет")
        self.mng.getFile("Киiвтеплоенерго")
        self.mng.getFile("Звiт_Рестр")
        self.mng.getFile("90%ТКЕ_ПСО")

        self.mng.deleteUnCalledFiles()               
        self.mng.allFromXlsToXlsx()
        #self.mng.printAllFiles()
        
        try:
            self.todayTKE = self.mng.getFile("Новый отчет", extension=".xlsx")
            self.yesterdayTKE = self.mng.getFile("90%ТКЕ_ПСО", extension=".xlsx")
            self.kyivEnergoPas = self.mng.getFile("Киiвтеплоенерго", extension=".xlsx")
            self.restructurization1730 = self.mng.getFile("Звiт_Рестр", extension=".xlsx")
            
            if self.mng.getNumberOfFiles() != self.numberOfFilesToStart:
                    raise AttributeError
        except AttributeError:
            print("Не хватает файлов для работы. Проверьте директорию " + str(path))
            msg = """Файлы, нужные для работы: 
            1. 90%ТКЕ_ПСО_... в пропуске обяательно должна стоять сегодняшняя дата. 
                            Например: На календаре 16.09.2019, название файла,
                            в таком случае: 90%ТКЕ_ПСО_16 или 90%ТКЕ_ПСО16 и т.д.
            2. 90%ТКЕ_ПСО_... в пропуске стоит дата вчерашнего дня. Это файл, с которым
                            будет сравниваться список предприятий (новички).
                            По аналогии, название может быть: 90%ТКЕ_ПСО_15
            Итого: 2 экселевских файлов
            После исправления запустите программу заново. Сейчас программа завершит работу
            Нажмите любую клавишу а затем Enter
            """
            print(msg)
            input()
            exit()

    def deleteFiles(self, programmIsDone=True):
        """Deletes all created files with .xlsx extension
        """
        # If programm has daone its work then close files
        if programmIsDone == True:
            try:
                self.yesterdayTKE.close()
                self.kyivEnergoPas.close()
                self.restructurization1730.close()
            except:
                print("Программа не смогла закрыть экселевские файлы")
        self.mng.deleteClosedFiles()
        return

    
        
    def run(self):
        """
        """
        #self.yesterdayTKE.open()
        self.copyColumn()

        fileNameForRead = "forRead"
        self.mng.createDuplicate(self.todayTKE, fileNameForRead)
        #Read from/write to block
        # Write to
        self.todayTKE.open(data_only=False)
        # self.todayTKE.save(self.todayTKE.pathToFile, fileNameForRead)
        todayWs = self.todayTKE.getWs("Sheet1")

        # Read from
        fileNameForRead+=".xlsx"
        todayTkeWithData = self.mng.addFileByPath(self.todayTKE.pathToFile, 
                            fileNameForRead, returnFile=True)
        todayTkeWithData.open(data_only=True)
        # todayTkeWithData.save(self.todayTKE.pathToFile, "rgrrrrr") #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        todayWsData = todayTkeWithData.getWs("Sheet1")
        self.mng.printAllFiles()

        numberOfRows = todayWs.max_row
        # Set 1 and 'договир э' to those companies who have a restructurization contract
        self.restructurization1730.open()
        rangeIter = "O12" + ":" + "O" + str(numberOfRows)
        for cells in todayWsData[rangeIter]:
            for cell in cells:
                if cell.value != "" and cell.value != None:
                    summary = self.restructurization(todayWsData, cell.column, cell.row)
                    column = openpyxl.utils.column_index_from_string("AM")
                    if summary != None:
                        todayWs.cell(column=column, row=cell.row).value = summary
                    elif summary == None:
                        todayWs.cell(column=column, row=cell.row).value = str("договір є")
                    # This check needs for empty cell not to be filled
                    column = openpyxl.utils.column_index_from_string("AT")
                    if todayWsData.cell(column=column, row=cell.row).value != "":
                        todayWs.cell(column=column, row=cell.row).value = 1
        # Transfer data from 'поточний лимит' to 'попередний лимит'
        rangeIter1 = "BO10" + ":" + "BU" + str(numberOfRows)
        rangeIter2 = "BB10" + ":" + "BH" + str(numberOfRows)
        for cells1, cells2 in zip(todayWsData[rangeIter1], todayWsData[rangeIter2]):
            for cell1, cell2 in zip(cells1, cells2):
                if cell1.row == cell2.row:
                    todayWs.cell(column=cell2.column, row=cell2.row).value = cell1.value

        # Set 'план э' to those rows wich have 0`s in both columns with conditions
        # Check the range
        list1 = todayTkeWithData.getListOfCellsByCriteria(0, "AS")  
        list2 = todayTkeWithData.getListOfCellsByCriteria(0, "AT")

        for cell1 in list1:
            if cell1.value == 0:
                for cell2 in list2:
                    if cell2.value == 0 and cell1.row == cell2.row:
                        column = openpyxl.utils.column_index_from_string("AU")
                        cellValueCheck = todayWsData.cell(column=column, row=cell1.row).value
                        if cellValueCheck == 0 or cellValueCheck == None or cellValueCheck == "":
                            continue
                        else:
                            column = openpyxl.utils.column_index_from_string("AU")
                            todayWs.cell(column=column, row=cell1.row).value = str("план є")
                            todayWs.cell(column=column+1, row=cell1.row).value = ""
                            todayWs.cell(column=column+2, row=cell1.row).value = ""
                            todayWs.cell(column=column+3, row=cell1.row).value = ""
                            todayWs.cell(column=column+4, row=cell1.row).value = ""
                            todayWs.cell(column=column+5, row=cell1.row).value = ""
                            todayWs.cell(column=column+6, row=cell1.row).value = ""

        # Find the difference between columns with 'план на декаду' and 'поточний лимит'
        for row in range(10, numberOfRows):
            column = openpyxl.utils.column_index_from_string("AU")
            cellValueCheck = todayWsData.cell(column=column, row=row).value
            if cellValueCheck == "план є" or cellValueCheck == None:
                continue
            else:
                column1 = openpyxl.utils.column_index_from_string("BO")
                value1 = todayWsData.cell(column=column1, row=row).value
                value2 = todayWsData.cell(column=column, row=row).value
                dx = value1 - value2
                if dx > 1e-6 or dx < -1e-6:
                    column = openpyxl.utils.column_index_from_string("BW")
                    todayWs.cell(column=column, row=row).value = dx

        self.kyivEnergoMoney(todayTkeWithData)

        self.todayTKE.save(self.todayTKE.pathToFile, "nn")
        # self.todayTKE.unmerge()
        return

    def copyColumn(self):

        # TODO: Make check for number of rows
        tmpTodayTKE = self.todayTKE
        tmpYesterdayTKE = self.yesterdayTKE

        self.todayTKE = self.mng.addFileByPath(self.todayTKE.pathToFile, 
                        self.todayTKE.fileName, returnFile=True, defaultParser=False, openBy=1)
        self.yesterdayTKE = self.mng.addFileByPath(self.yesterdayTKE.pathToFile, 
                        self.yesterdayTKE.fileName, returnFile=True, defaultParser=False, openBy=1)
        self.todayTKE.open()
        self.yesterdayTKE.open()

        todayWs = self.todayTKE.getWs("Sheet1")
        yestWs = self.yesterdayTKE.getWs("Sheet1")

        self.todayTKE.incertColumn("AS")
        todayWs.Range("AS1:AS2").EntireColumn.Unmerge()
        yestWs.Range("AS1:AS2").EntireColumn.Unmerge()
        yestWs.Range("AS1:AS2").EntireColumn.Copy()
        todayWs.Paste(todayWs.Range("AS1:AS2"))

        self.todayTKE.save(self.todayTKE.pathToFile, self.todayTKE.fileNameWithoutExtension)
        self.yesterdayTKE.save(self.yesterdayTKE.pathToFile, self.yesterdayTKE.fileNameWithoutExtension)
        self.todayTKE.close()
        self.mng.removeUnCalledFiles()
        #self.todayTKE = self.mng.addFileByPath(self.todayTKE.pathToFile, self.todayTKE.fileName)
        #self.yesterdayTKE = self.mng.addFileByPath(self.yesterdayTKE.pathToFile, self.yesterdayTKE.fileName)
        self.todayTKE = tmpTodayTKE
        self.yesterdayTKE = tmpYesterdayTKE

        return

    def restructurization(self, ws, column: int, row: int):
        EDRPOU = ws.cell(column=column+1, row=row).value
        try:
            row = self.restructurization1730.getFirstCellByCriteria(EDRPOU, "D").row-1
        except AttributeError:
            print("В списках договоров реструктуризации " + \
                        "1730 не найдено предприятие, с кодом ЕДРПОУ", EDRPOU)
            return None
        wsRestr = self.restructurization1730.getWs("Sheet1")
        overpaymentColumn = openpyxl.utils.column_index_from_string("U")
        debtColumn = openpyxl.utils.column_index_from_string("V")
        overpayment = wsRestr.cell(column=overpaymentColumn, row=row).value
        debt = wsRestr.cell(column=debtColumn, row=row).value
        summary = overpayment + debt
        if summary > 0:
            return summary
        if summary <= 0:
            return None

    def kyivEnergoMoney(self, dataFile):
        try:
            self.kyivEnergoPas.open()
            kyivWs = self.kyivEnergoPas.getWs("Sheet1")
            
            headerColumn = openpyxl.utils.column_index_from_string(str("A"))
            contractColumn = openpyxl.utils.column_index_from_string(str("B"))
            moneyColumn = openpyxl.utils.column_index_from_string(str("I"))
            
            listOfHeaders = self.kyivEnergoPas.getListOfCellsByCriteria("Період", "A")
            row = listOfHeaders[1].row
            
            money = 0
            while True:
                row += 1
                header = kyivWs.cell(column=headerColumn, row=row).value
                if header == "" or header == None:
                    break
                if "рік" in header:
                    contract = kyivWs.cell(column=contractColumn, row=row).value
                    if "РЗ" not in contract:
                        money += kyivWs.cell(column=moneyColumn, row=row).value
            
            self.kyivEnergoPas.close()
        except:
            money = 0
            print("Проблема с подсчетом оплаты Київтеплоенерго КП ВО")
        
        try:
            kyivEnergoRow = dataFile.getFirstCellByCriteria("Київтеплоенерго " + \
                            "КП ВО Київради (КМДА)", "R").row
            ws = self.todayTKE.getWs("Sheet1")

            # Payment column value
            column=openpyxl.utils.column_index_from_string(str("AQ"))
            ws.cell(column=column, row=kyivEnergoRow).value = money
            # For all contracts column value
            column=openpyxl.utils.column_index_from_string(str("AE"))
            column1=openpyxl.utils.column_index_from_string(str("AF"))
            column2=openpyxl.utils.column_index_from_string(str("AD"))
            ws.cell(column=column, row=kyivEnergoRow).value = \
                            ws.cell(column=column1, row=kyivEnergoRow).value - \
                            ws.cell(column=column2, row=kyivEnergoRow).value - money
        except:
            print("Программа не смогла внести данные о задолженности Київтеплоенерго КП ВО")
        
        return 
    

    # def hideColumns(self):
    #     for column in range(1, self.todayTKE.ws.max_column):
    #         if column < len(self.listOfHiddenColumns):
    #             if self.listOfHiddenColumns[column] == True:
    #                 self.todayTKE.ws.column_dimensions[openpyxl.utils.get_column_letter(
    #                     column)].hidden = True
    #     return

