import os
import datetime
from copy import copy

import openpyxl
import win32com.client


from src.manager import *


class TKE:

    def __init__(self, dir: str):
        self.mng = Manager()
        self.numberOfFilesToStart = 4
        self.checkIfDirectoryIsReady(dir)

    def checkIfDirectoryIsReady(self, path: str):
        self.mng.setWorkDir(os.path.abspath(path))
        self.mng.addFilesInDir()

        self.mng.getFile("Новый отчет")
        self.mng.getFile("Киiвтеплоенерго")
        self.mng.getFile("Звiт_Рестр")
        self.mng.getFile("90%ТКЕ_ПСО")

        self.mng.deleteUnCalledFiles()               
        self.mng.allFromXlsToXlsx()
        self.mng.printAllFiles()
        
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

    def CopyColumn(self):

        # TODO: Make check for number of rows

        self.todayTKE = self.mng.addFileByPath(self.todayTKE.pathToFile, self.todayTKE.fileName, defaultParser=False, openBy=1)
        self.yesterdayTKE = self.mng.addFileByPath(self.yesterdayTKE.pathToFile, self.yesterdayTKE.fileName, defaultParser=False, openBy=1)
        self.mng.printAllFiles()
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
        self.todayTKE = self.mng.addFileByPath(self.todayTKE.pathToFile, self.todayTKE.fileName)
        self.yesterdayTKE = self.mng.addFileByPath(self.yesterdayTKE.pathToFile, self.yesterdayTKE.fileName)

    def run(self):
        """
        """
        #self.yesterdayTKE.open()
        self.CopyColumn()
        self.todayTKE.open()
        todayWs = self.todayTKE.getWs("Sheet1")
        numberOfRows = todayWs.max_row
        # Set 1 and 'договир э' to those companies who have a restructurization contract
        rangeIter = "O12" + ":" + "O" + str(numberOfRows)
        for cells in todayWs[rangeIter]:
            for cell in cells:
                if cell.value != "" and cell.value != None:
                    todayWs[str("AM")+str(cell.row)
                                     ] = str("договір є")
                    # This check needs for empty cell not to be filled
                    if todayWs[str("AT")+str(cell.row)] != "":
                        todayWs[str("AT")+str(cell.row)] = 1

        # Transfer data from 'поточний лимит' to 'попередний лимит'
        rangeIter1 = "BO10" + ":" + "BU" + str(numberOfRows)
        rangeIter2 = "BB10" + ":" + "BH" + str(numberOfRows)
        for cells1, cells2 in zip(todayWs[rangeIter1], todayWs[rangeIter2]):
            for cell1, cell2 in zip(cells1, cells2):
                if cell1.row == cell2.row:
                    todayWs.cell(
                        column=cell2.column, row=cell2.row, value=cell1.value)

        # Set 'план э' to those rows wich have 0`s in both columns with conditions
        # Check the range
        list1 = self.todayTKE.getListOfCellsByCriteria(0, "AS")  
        list2 = self.todayTKE.getListOfCellsByCriteria(0, "AT")

        for cell1 in list1:
            if cell1.value == 0:
                for cell2 in list2:
                    if cell2.value == 0 and cell1.row == cell2.row:
                        todayWs[str(
                            "AU")+str(cell1.row)] = str("план є")
                        todayWs[str("AV")+str(cell1.row)] = ""
                        todayWs[str("AW")+str(cell1.row)] = ""
                        todayWs[str("AX")+str(cell1.row)] = ""
                        todayWs[str("AY")+str(cell1.row)] = ""
                        todayWs[str("AZ")+str(cell1.row)] = ""
                        todayWs[str("BA")+str(cell1.row)] = ""

        # Find the difference between columns with 'план на декаду' and 'поточний лимит'
        for row in range(10, numberOfRows):
            if todayWs[str("AU")+str(row)].value != "план є" and todayWs[str("AU")+str(row)].value != None:
                dx = todayWs[str(
                    "BO")+str(row)].value-todayWs[str("AU")+str(row)].value
                if dx > 1e-1 or dx < -1e-4:                                                         # Check range!!!
                    todayWs.cell(column=openpyxl.utils.column_index_from_string(str("BW")),
                                          row=row,
                                          value=dx
                                          )
        self.todayTKE.save(self.todayTKE.pathToFile, "nn")
        # self.todayTKE.unmerge()
        return

    

    # def hideColumns(self):
    #     for column in range(1, self.todayTKE.ws.max_column):
    #         if column < len(self.listOfHiddenColumns):
    #             if self.listOfHiddenColumns[column] == True:
    #                 self.todayTKE.ws.column_dimensions[openpyxl.utils.get_column_letter(
    #                     column)].hidden = True
    #     return

