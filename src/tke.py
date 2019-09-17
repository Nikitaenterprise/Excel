import datetime
import os
from copy import copy

import openpyxl
import win32com.client

from src.excel import ExcelBook, hasNumbers


class TKE:

    def __init__(self, dir: str):
        self.checkIfDirectoryIsReady(dir)

    def checkIfDirectoryIsReady(self, path: str):
        numberOfFiles = self.scanDirectory(path)
        # Check the dir for needed files
        while True:
            if numberOfFiles == 2:
                break
            if numberOfFiles > 2:
                print("Слишком много экселевских файлов в папке")
                print("Должно быть ровно" + str(numberOfFiles))
                print("Программа пробует удалить ненужные")
                self.deleteFiles(False)
            numberOfFiles = self.scanDirectory(path)

        try:
            self.todayTKE
            self.yesterdayTKE
        except AttributeError:
            print("Не хватает файлов для работы. Проверьте директорию " + str(path))
            print(self.instructionMessage())
            input()
            exit()

    def scanDirectory(self, path: str):
        """Scans the directory with os.walk() for excel files
        and set class excel book variables for folowing work
        """
        print(os.path.abspath(path))
        numberOfFiles = 0
        # r=root, d=directories, f = files
        for r, d, f in os.walk(path):
            for file in f:
                if ".xls" in file or ".xlsx" in file:
                    numberOfFiles += 1
                    pathToFile = path+"\\"+file
                    if "90%ТКЕ" in file:
                        if hasNumbers(file):
                            if str(datetime.datetime.today().day) in file:
                                self.todayTKE = ExcelBook(
                                    pathToFile, read=False)
                            elif str(datetime.datetime.today().day - 1) in file:
                                self.yesterdayTKE = ExcelBook(
                                    pathToFile, read=False)
                            else:
                                print(
                                    "Будьте осторожны, программа использует файл с деньгами с неправильной датой")
                                self.yesterdayTKE = ExcelBook(
                                    pathToFile, read=False)
        return numberOfFiles

    def instructionMessage(self):
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
        return msg

    def deleteFiles(self, programmIsDone=True):
        """Deletes all created files with .xlsx extension
        """
        # If programm has daone its work then close files
        if programmIsDone == True:
            try:
                self.closeFiles()
            except:
                print("Программа не смогла закрыть экселевские файлы")

        fileNameWithPathWithoutExtensionTodayTKE = os.path.splitext(
            self.todayTKE.fileNameWithPath)[0]
        fileNameWithPathWithoutExtensionYesterdayTKE = os.path.splitext(
            self.yesterdayTKE.fileNameWithPath)[0]

        numberOfDeletedFiles = 0
        try:
            os.remove(fileNameWithPathWithoutExtensionTodayTKE + ".xlsx")
        except FileNotFoundError:
            numberOfDeletedFiles += 1
        try:
            os.remove(fileNameWithPathWithoutExtensionYesterdayTKE + ".xlsx")
        except FileNotFoundError:
            numberOfDeletedFiles += 1
        return

    def run(self):
        """
        """
        self.todayTKE.readExcelFile()
        self.yesterdayTKE.readExcelFile()

        self.todayTKE.incertColumnWithPyWin("AS")
        excelAppYesterday, wbYesterday = self.yesterdayTKE.readFileWithPyWin()
        excelAppToday, wbToday = self.todayTKE.readFileWithPyWin()

        excelYesterday = win32com.client.Dispatch("Excel.Application")
        excelYesterday.Visible = False
        yesterdayTkePyWin = excelYesterday.Workbooks.Open(
            self.yesterdayTKE.fileNameWithPath)
        sheetYesterday = workbookYesterday.Worksheets(1)
        # sheetYesterday.Range("AS1:AS"+str(self.yesterdayTKE.ws.max_row)).Unmerge()
        # sheet.Range("AS1:AS"+str(numberOfRows)).Unmerge()
        yesterdayColumn = sheetYesterday.Range(
            "AS1:AS"+str(self.yesterdayTKE.ws.max_row)).Copy()
        sheet.Paste(sheet.Range("AS1:AS"+str(numberOfRows)))
        workbookYesterday.Close()

        self.todayTKE.readExcelFile()
        numberOfRows = self.todayTKE.ws.max_row
        # Set 1 and 'договир э' to those companies who have a restructurization contract
        rangeIter = "O12" + ":" + "O" + str(numberOfRows)
        for cells in self.todayTKE.ws[rangeIter]:
            for cell in cells:
                if cell.value != "" and cell.value != None:
                    self.todayTKE.ws[str("AM")+str(cell.row)
                                     ] = str("договір є")
                    # This check needs for empty cell not to be filled
                    if self.todayTKE.ws[str("AT")+str(cell.row)] != "":
                        self.todayTKE.ws[str("AT")+str(cell.row)] = 1

        # Opens ysterday workbook and copies one specific column to current
        # workbook shifting other columns
        # fillColor = openpyxl.styles.PatternFill(start_color="cdffcd",
        #                                        fill_type="solid"
        #                                        ) #Set fill color green
        # fillBlue = openpyxl.styles.PatternFill(start_color="9acdff",
        #                                        fill_type="solid"
        #                                        ) #Set fill color blue

        # font = openpyxl.styles.Font(name="Arial",sz=9)      # Set font
        # border = openpyxl.styles.Border(left=openpyxl.styles.Side(border_style="thin",
        #                                color='000000'),
        #                                right=openpyxl.styles.Side(border_style="thin",
        #                                color='000000'),
        #                                top=openpyxl.styles.Side(border_style="thin",
        #                                color='000000'),
        #                                bottom=openpyxl.styles.Side(border_style="thin",
        #                                color='000000')
        #                                )                       # Set border

        # excel = win32com.client.Dispatch("Excel.Application")
        # excel.Visible = False
        # workbook = excel.Workbooks.Open("C:\Code\Python\Excel\ggg.xlsx")
        # sheet = workbook.Worksheets(1)
        # rangeObj = sheet.Range("AS1:AS2")
        # rangeObj.EntireColumn.Insert()
        # del rangeObj

        wbFromYesterday = ExcelBook("./second.xlsx")

        excelYesterday = win32com.client.Dispatch("Excel.Application")
        excelYesterday.Visible = False
        workbookYesterday = excelYesterday.Workbooks.Open(
            "C:\Code\Python\Excel\second.xlsx")  # Right path
        sheetYesterday = workbookYesterday.Worksheets(1)
        sheetYesterday.Range(
            "AN1:AN"+str(wbFromYesterday.ws.max_row)).Unmerge()
        sheet.Range("AS1:AS"+str(numberOfRows)).Unmerge()
        rangeObjYesterday = sheetYesterday.Range(
            "AN1:AN"+str(wbFromYesterday.ws.max_row)).Copy()
        sheet.Paste(sheet.Range("AS1:AS"+str(numberOfRows)))
        workbookYesterday.Close()
        workbook.Save()
        workbook.Close()
        excelYesterday.Quit()
        excel.Quit()

        columnNumber = openpyxl.utils.column_index_from_string("AS")
        if numberOfRows == wbFromYesterday.ws.max_row:  # TODO make the right check
            rangeIter = "AS1" + ":" + "BU" + str(numberOfRows)
            self.todayTKE.ws.move_range(rangeIter, rows=0, cols=1)
            # for cell in wbFromYesterday.ws["AN"]:
            #     if cell.value == "" or cell.value == None:
            #         self.todayTKE.ws.cell(row = cell.row,
            #                     column = columnNumber,
            #                     ).fill = fillBlue
            #     else:
            #         self.todayTKE.ws.cell(row = cell.row,
            #                         column = columnNumber,
            #                         value = cell.value
            #                         )
            #         self.todayTKE.ws.cell(row = cell.row,
            #                         column = columnNumber,
            #                         ).font = font
            #         self.todayTKE.ws.cell(row = cell.row,
            #                         column = columnNumber,
            #                         ).border = border
            #         self.todayTKE.ws.cell(row = cell.row,
            #                         column = columnNumber,
            #                         ).fill = fillColor
        else:
            raise Exception("Different number of rows in both docs\n. The first has: {}".format(
                self.todayTKE.ws.max_row))

        # Copy names of hidden columns
        self.listOfHiddenColumns = []
        for column in range(1, wbFromYesterday.ws.max_column):
            self.listOfHiddenColumns.append(
                wbFromYesterday.ws.column_dimensions[openpyxl.utils.get_column_letter(column)].hidden)

        listOfRangedCells = wbFromYesterday.ws.merged_cells.ranges
        wbFromYesterday.close()

        # Transfer data from 'поточний лимит' to 'попередний лимит'
        rangeIter1 = "BO10" + ":" + "BU" + str(numberOfRows)
        rangeIter2 = "BB10" + ":" + "BH" + str(numberOfRows)
        for cells1, cells2 in zip(self.todayTKE.ws[rangeIter1], self.todayTKE.ws[rangeIter2]):
            for cell1, cell2 in zip(cells1, cells2):
                if cell1.row == cell2.row:
                    self.todayTKE.ws.cell(
                        column=cell2.column, row=cell2.row, value=cell1.value)

        # Multiply cells in column 'план на декаду' by 3
        rangeIter = "AU10" + ":" + "BA" + str(numberOfRows)
        for cells in self.todayTKE.ws[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    self.todayTKE.ws.cell(
                        column=cell.column, row=cell.row, value=cell.value*3)

        # Set 'план э' to those rows wich have 0`s in both columns with conditions
        list1 = self.todayTKE.getListOfCellsWithCriteria(
            "AS", 0)             # Check the range
        list2 = self.todayTKE.getListOfCellsWithCriteria(
            "AT", 0)             # Check the range

        for cell1 in list1:
            if cell1.value == 0:
                for cell2 in list2:
                    if cell2.value == 0 and cell1.row == cell2.row:
                        self.todayTKE.ws[str(
                            "AU")+str(cell1.row)] = str("план є")
                        self.todayTKE.ws[str("AV")+str(cell1.row)] = ""
                        self.todayTKE.ws[str("AW")+str(cell1.row)] = ""
                        self.todayTKE.ws[str("AX")+str(cell1.row)] = ""
                        self.todayTKE.ws[str("AY")+str(cell1.row)] = ""
                        self.todayTKE.ws[str("AZ")+str(cell1.row)] = ""
                        self.todayTKE.ws[str("BA")+str(cell1.row)] = ""

        # Find the difference between columns with 'план на декаду' and 'поточний лимит'
        for row in range(10, numberOfRows):
            if self.todayTKE.ws[str("AU")+str(row)].value != "план є" and self.todayTKE.ws[str("AU")+str(row)].value != None:
                dx = self.todayTKE.ws[str(
                    "BO")+str(row)].value-self.todayTKE.ws[str("AU")+str(row)].value
                if dx > 1E-4 or dx < -1E-4:                                                         # Check range!!!
                    self.todayTKE.ws.cell(column=openpyxl.utils.column_index_from_string(str("BW")),
                                          row=row,
                                          value=dx
                                          )

        # self.todayTKE.unmerge()
        return

    def hideColumns(self):
        for column in range(1, self.todayTKE.ws.max_column):
            if column < len(self.listOfHiddenColumns):
                if self.listOfHiddenColumns[column] == True:
                    self.todayTKE.ws.column_dimensions[openpyxl.utils.get_column_letter(
                        column)].hidden = True
        return


# first = TKE("./first.xlsx")
# first.start()
# first.hideColumns()
# first.save("./out.xlsx")
