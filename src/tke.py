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
        self.mng.addFileByPath(self.todayTKE.pathToFile, 
                            self.generateName() + ".xlsx")
        self.mng.deleteClosedFiles()
        return

    def run(self):
        self.copyColumn()
        self.mainCalculations()
        name = self.generateName()
        self.todayTKE.save(self.todayTKE.pathToFile, name, extension=".xls")
        self.deleteFiles()
        
    def mainCalculations(self):
        """
        """
        fileNameForRead = "forRead"
        self.mng.createDuplicate(self.todayTKE, fileNameForRead)
        #Read from/write to block
        # Write to
        self.todayTKE.open(data_only=False)
        todayWs = self.todayTKE.getWs("Sheet1")

        # Read from
        fileNameForRead+=".xlsx"
        todayTkeWithData = self.mng.addFileByPath(self.todayTKE.pathToFile, 
                            fileNameForRead, returnFile=True)
        todayTkeWithData.open(data_only=True)
        todayWsData = todayTkeWithData.getWs("Sheet1")

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
            if cellValueCheck == None:
                continue
            else:
                columnAS = openpyxl.utils.column_index_from_string("AS")
                columnAT = openpyxl.utils.column_index_from_string("AT")
                cellValueCheck1 = todayWsData.cell(column=columnAS, row=row).value
                cellValueCheck2 = todayWsData.cell(column=columnAT, row=row).value
                if cellValueCheck1 == 0 and cellValueCheck2 == 0:
                    continue
                elif cellValueCheck1 == None and cellValueCheck2 == None:
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
        todayTkeWithData.close()
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

        column = openpyxl.utils.column_index_from_string("R")
        for row in range(1, todayWs.UsedRange.Rows.Count):
            value1 = todayWs.Cells(row, column).Value
            value2 = yestWs.Cells(row, column).Value
            if value1 != value2:
                self.yesterdayTKE.insertRow(str(row))
                for column1 in range(1, yestWs.UsedRange.Columns.Count):
                    yestWs.Cells(row, column1).Value = todayWs.Cells(row, column1).Value
                    print(yestWs.Cells(row, column1).Value)

        self.todayTKE.insertColumn("AS")
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
                        "1730 не найдено предприятие с кодом ЕДРПОУ", EDRPOU)
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

    def smilaTeplo(self, dataFile):
        try:
            smilaRow = dataFile.getFirstCellByCriteria("Смілакомунтеплоенерго КП", "R").row
            ws = self.todayTKE.getWs("Sheet1")
            # For all contracts column value
            column=openpyxl.utils.column_index_from_string(str("AE"))
            column1=openpyxl.utils.column_index_from_string(str("AF"))
            # Payment column value
            column2=openpyxl.utils.column_index_from_string(str("AQ"))
            ws.cell(column=column, row=smilaRow).value = \
                            ws.cell(column=column1, row=smilaRow).value - \
                            ws.cell(column=column2, row=smilaRow).value
        except:
            print("Программа не смогла внести данные о задолженности Смілакомунтеплоенерго КП")


    def generateName(self):
        day = datetime.datetime.today().day
        month = datetime.datetime.today().month
        year = datetime.datetime.today().year
        monthInRussian = [r"январь", r"февраль", r"март", 
                            r"апрель", r"май", r"июнь", 
                            r"июль", r"август", r"сентябрь", 
                            r"октябрь", r"ноябрь", r"декабрь"]
        fileName = "90%ТКЕ_ПСО_" + monthInRussian[month-1]
        fileName += "(" + str(day) + "."
        if month < 10:
            fileName += "0" + str(month) + "."
        elif month >= 10:
            fileName += str(month) + "."
        fileName += str(year) + ")"
        return fileName
    

    # def hideColumns(self):
    #     for column in range(1, self.todayTKE.ws.max_column):
    #         if column < len(self.listOfHiddenColumns):
    #             if self.listOfHiddenColumns[column] == True:
    #                 self.todayTKE.ws.column_dimensions[openpyxl.utils.get_column_letter(
    #                     column)].hidden = True
    #     return

