import os


import openpyxl
import win32com


def hasNumbers(inputString: str):
        """Checks string for containing numbers
        returns True if string has at least one digit
        """
        return any(char.isdigit() for char in inputString)

class ExcelBook():

    def __init__(self, name: str, data_only=True, read=True, worksheet: int=0, keep_vba=True):
        self.fileNameWithPath = name
        self.data_only = data_only
        self.worksheetNumberInBook = worksheet
        self.keep_vba = keep_vba
        if read == True:
            self.readExcelFile()

    def readExcelFile(self):
        """Reads file after creation of class instance by default.
        If this function wasn`t called after initialization of class instance
        (for example if you dont want to allocate a lot of memory by opening
        many excel books, but you want to initialize a class instances)
        then it can be called after.
        """
        extension = os.path.splitext(self.fileNameWithPath)[1] # may be .xls or .xlsx
        if extension == ".xls":
            self.reSaveFromXlsToXlsx(self.fileNameWithPath)
            extension = os.path.splitext(self.fileNameWithPath)[1] # reinitialize extension
        if extension == ".xlsx":
            self.wb = openpyxl.load_workbook(self.fileNameWithPath, data_only=self.data_only, keep_vba=self.keep_vba)
            self.ws = self.wb[self.wb.sheetnames[self.worksheetNumberInBook]]
        return True

    def readFileWithPyWin(self):
        """
        """
        extension = os.path.splitext(self.fileNameWithPath)[1] # may be .xls or .xlsx
        if extension == ".xls" or extension == ".xlsx":
            excelApp = win32com.client.Dispatch("Excel.Application")
            excelApp.Visible = False

            #try:
            wb = excelApp.Workbooks.Open(os.path.abspath(self.fileNameWithPath))
            # except:
            excelApp.Quit()
            #     print("Программа не может открыть файл " + self.fileNameWithPath)
            #     raise FileNotFoundError
        return wb

    def reSaveFromXlsToXlsx(self, name: str):
        """Opens file in .xls format and saves it 
        in .xlsx format using pyWin32
        
        Keyword arguments:
        name -- name of file that will be changed
        """
        newName = os.path.splitext(name)[0]
        newName += str(".xlsx")
        excelApp = win32com.client.Dispatch("Excel.Application")
        excelApp.Visible = False

        try:
            wb = excelApp.Workbooks.Open(os.path.abspath(name))
        except:
            excelApp.Quit()
            print("Программа не может открыть файл " + name)
            raise FileNotFoundError

        xlsx = 51 # Code for xslx format
        try:
            wb.SaveAs(os.path.abspath(newName), FileFormat=xlsx)
            excelApp.Quit()
        except:
            print("Программа не может сохранить файл " + newName)
            raise WindowsError
        
        self.fileNameWithPath = newName
        return

    def save(self, name: str):
        """Closes file and save it to project root dir
        with a given name

        Keyword arguments:
        name -- name of saved file
        """
        self.wb.save(name)
        self.wb.close()
        return
    
    def close(self):
        """Closes file without saving
        """
        self.wb.close()
        return 

    def setCellsInColumnByRowCoord(self, row: int, column: str, value):
        """Finds values in one column by row coordinate
        and then set value of that cell

        Keyword arguments:
        row -- row number
        column -- column in which the searh would happend
        value -- this would be set to the cell
        """
        for cells in self.ws[column]:
            for cell in cells:
                if cell.row == row:
                    cell.value = value
        return

    def setCellsInRowByColumnCoord(self, row: int, column: str, value):
        """Finds values in one row by column coordinate
        and then set value of that cell

        Keyword arguments:
        row -- row number
        column -- column in which the searh would happend
        value -- this would be set to the cell
        """
        for cells in self.ws[str(row)]:
            for cell in cells:
                if cell.column == openpyxl.utils.column_index_from_string(column):
                    cell.value = value
        return

    def getFirstCellByCriteria(self, criteria, range: str = None):
        """Finds first cell by searchin in whole sheet or in 
        some range the target criteria. Similar to Excel 
        function VLOOKUP (ВПР)

        Keyword arguments:
        value -- searching value (str, int, ...)
        range -- search range (I22:J22, or I), by default set to None
                    so it search in whole sheet
        """
        if range == None:
            diapason = self.ws
            for cells in diapason:
                for cell in cells:
                    if cell.value == criteria:
                        return cell
        elif range != None:
            diapason = self.ws[range]
            if  hasNumbers(range) == True:
                for cells in diapason:
                    for cell in cells:
                        if cell.value == criteria:
                            return cell
            elif hasNumbers(range) == False:
                for cell in diapason:
                    if cell.value == criteria:
                        return cell
        return None

    def getListOfCellsByCriteria(self, criteria, range: str):
        """Finds list of cells with values equal to criteria by
        searching in some range or in whole sheet

        Keyword arguments:
        criteria -- search criteria in cells values
        range -- search range (I22:J22, I, ...), by default set to None
                    so it search in whole sheet
        """
        listOfCells = []

        if  hasNumbers(range) == True:
            for cells in self.ws[range]:
                for cell in cells:
                    if cell.value == criteria:
                        listOfCells.append(cell)
        elif hasNumbers(range) == False:
            for cell in self.ws[range]:
                if cell.value == criteria:
                    listOfCells.append(cell)
        return listOfCells

    def checkRowForRegion(self, row: int, column: int):
        """In excel files sometimes it can be observed that some rows
        contains sum of data of other rows below or under them.
        Such rows often contain name of 'область' on their left column
        Under them there are companies in this district.
        This function checks the row for containing name of district
        in specific column

        TODO

        Keyword arguments:
        range
        """
        return

    def unmerge(self):
        for range in self.ws.merged_cells.ranges:
            rangeList = list(range.bounds)
            minCol = rangeList[0]
            minRow = rangeList[1]
            maxCol = rangeList[2]
            maxRow = rangeList[3]
            self.ws.unmerge_cells(start_row=minRow,
                                start_column=minCol,
                                end_row=maxRow,
                                end_column=maxCol
                                )
        return


    def merge(self, range: str):
        start = range.split(":")[0]
        end = range.split(":")[1]
        minRow = openpyxl.utils.coordinate_to_tuple(start)[0]
        minCol = openpyxl.utils.coordinate_to_tuple(start)[1]
        maxRow = openpyxl.utils.coordinate_to_tuple(end)[0]
        maxCol = openpyxl.utils.coordinate_to_tuple(end)[1]
        self.ws.merge_cells(start_row=minRow,
                            start_column=minCol,
                            end_row=maxRow,
                            end_column=maxCol
                            )
        return

    def mergeByTuple(self, rangeList: list):
        for range in rangeList:
            coord = list(range.bounds)
            rangeStr = str(openpyxl.utils.get_column_letter(coord[0])) + str(coord[1]) + ":" + str(openpyxl.utils.get_column_letter(coord[2])) + str(coord[3])
            self.merge(rangeStr)
        return

    
    
    # def addHeader(self, range: str)
    #     """Adds header class into this class

    #     Keyword arguments:
    #     range -- header diapasone (I22:J22, I)
    #     """
    #     self.header = Header()
    #     return
class Header():

    def __init__(self, headerRange: str):
        self.header = headerRange
        split = headerRange.split(":")
        
        try:
            if len(split) < 2:
                raise Exception()
        except Exception:
            print("Header can`t be 1 cell " + headerRange)

        self.diapasone = self.excelWorksheet.ws[self.header]
        self.leftTopCoordinate = headerRange.split(":")[0]
        self.rightBotCoordinate = headerRange.split(":")[1]
        self.leftTopRow = openpyxl.utils.coordinate_to_tuple(self.leftTopCoordinate)[0]
        self.leftTopColumn = openpyxl.utils.coordinate_to_tuple(self.leftTopCoordinate)[1]
        self.rightBotRow = openpyxl.utils.coordinate_to_tuple(self.rightBotCoordinate)[0]
        self.rightBotColumn = openpyxl.utils.coordinate_to_tuple(self.rightBotCoordinate)[1]

    def getHeadersOfAllColumns(self):
        """
        """
        listOfHeaders = [[]]
        counter = 0
        for cells in self.diapasone:
            for cell in cells:
                if cell.value != None:
                    listOfHeaders[counter].append(cell.value)
                    listOfHeaders[counter].append(cell.column)
                    listOfHeaders[counter].append(cell.row)
                    counter += 1

    def findCellByStr(self, value: str):
        """Finds first cell by searchin in the header 
        the target value. Returns a list [row, column]

        Keyword argument:
        value -- searching string
        """
        for cells in self.ws[self.header]:
            for cell in cells:
                if cell.value == value:
                    return [cell.row, cell.column]

if __name__ == "__main__":
    print("I`m ExcelBook.py file")
