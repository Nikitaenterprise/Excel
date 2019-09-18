import os

import openpyxl
import win32com.client


def hasNumbers(inputString: str):
    """Checks string for containing numbers
    returns True if string has at least one digit
    """
    return any(char.isdigit() for char in inputString)


class File:

    def __init__(self, pathToFile: str, fileName: str):
        self.pathToFile = pathToFile
        self.fileName = fileName
        self.fileExtension = os.path.splitext(self.fileName)[1]
        self.fileNameWithoutExtension = os.path.splitext(self.fileName)[0]
        self.isOpened = False
        self.wasCalled = False

    def open(self):
        pass

    def close(self):
        pass

    def save(self, path: str, name: str, extension=".xlsx"):
        pass

    def getWb(self):
        pass
    
    def getWs(self):
        pass

class PyWin(File):

    def open(self):
        if self.isOpened == False:
            self.excelApp = win32com.client.Dispatch("Excel.Application")
            self.excelApp.Visible = False
            try:
                self.wb = self.excelApp.Workbooks.Open(
                    self.pathToFile+"\\"+self.fileName)
                self.isOpened = True
            except:
                self.isOpened = False
                self.excelApp.Quit()
                print("Программа не может открыть файл " + self.fileName)
                raise FileNotFoundError

    def getWb(self):
        if self.isOpened == True:
            return self.wb

    def getWs(self, wsName="", isActiveSheet=False):
        if self.isOpened == True:
            if isActiveSheet == False:
                return self.wb.Worksheets(wsName)
            elif isActiveSheet == True:
                return self.wb.ActiveSheet 

    def getApp(self):
        if self.isOpened == True:
            return self.excelApp

    def close(self):
        if self.isOpened == True:
            self.wb.Close()
            self.excelApp.Quit()
            self.isOpened = False

    def save(self, path: str, name: str, extension=".xlsx"):
        if self.isOpened == True:
            if extension == ".xlsx":
                # Code for xslx format
                fileFormat = 51
            elif extension == ".xls":
                # Code for xls format
                fileFormat = 56
            self.wb.SaveAs(path + "\\" + name + extension, FileFormat=fileFormat)

    def incertColumn(self, column: str):
        """Incerts column using pyWin. 
        Function incerts column to the right at first worksheet

        Keyword arguments:
        column -- name of column near what the column would be inserted
        """
        ws = self.getWs(isActiveSheet=True)
        # select column as range object
        if column.isdigit():
            openpyxl.utils.get_column_letter(column)
            newColumn = openpyxl.utils.get_column_letter(column)
            rangeObj = ws.Range(newColumn+str(1)+str(":")+newColumn+str(2))
        else:
            rangeObj = ws.Range(column+str(1)+str(":")+column+str(2))

        rangeObj.EntireColumn.Insert()
        self.save(self.pathToFile, self.fileName)

class OpenPyXl(File):
    def open(self, data_only=True, keep_vba=False):
        if self.isOpened == False:
            try:
                self.wb = openpyxl.load_workbook(self.pathToFile + "\\" + self.fileName,
                                                 data_only=data_only, keep_vba=keep_vba)
                self.isOpened = True
            except:
                self.isOpened = False
                print("Программа не может открыть файл " + self.fileName)
                raise FileNotFoundError

    def getWb(self):
        if self.isOpened == True:
            return self.wb

    def getWs(self, wsName="", isActiveSheet=False):
        if self.isOpened == True:
            if isActiveSheet == True or wsName == 0 or wsName == "":
                return self.wb[self.getWsNames()[0]]
            elif isActiveSheet == False:
                return self.wb[wsName]

    def getWsNames(self):
        if self.isOpened == True:
            return self.wb.sheetnames

    def close(self):
        if self.isOpened == True:
            self.wb.close()
            self.isOpened = False

    def save(self, path: str, name: str, extension=".xlsx"):
        if self.isOpened == True:
            self.wb.save(path+name)

    def setCellsInColumnByRowCoord(self, row: int, column: str, value, wsName=""):
        """Finds values in one column by row coordinate
        and then set value of that cell

        Keyword arguments:
        row -- row number
        column -- column in which the searh would happend
        value -- this would be set to the cell
        """
        for cells in self.getWs(wsName)[column]:
            for cell in cells:
                if cell.row == row:
                    cell.value = value
        return

    def setCellsInRowByColumnCoord(self, row: int, column: str, value, wsName=""):
        """Finds values in one row by column coordinate
        and then set value of that cell

        Keyword arguments:
        row -- row number
        column -- column in which the searh would happend
        value -- this would be set to the cell
        """
        for cells in self.getWs(wsName)[str(row)]:
            for cell in cells:
                if cell.column == openpyxl.utils.column_index_from_string(column):
                    cell.value = value
        return

    def getFirstCellByCriteria(self, criteria, range: str = None, wsName=""):
        """Finds first cell by searchin in whole sheet or in 
        some range the target criteria. Similar to Excel 
        function VLOOKUP (ВПР)

        Keyword arguments:
        value -- searching value (str, int, ...)
        range -- search range (I22:J22, or I), by default set to None
                    so it search in whole sheet
        """
        if range == None:
            diapason = self.getWs(wsName)
            for cells in diapason:
                for cell in cells:
                    if cell.value == criteria:
                        return cell
        elif range != None:
            diapason = self.getWs(wsName)[range]
            if hasNumbers(range) == True:
                for cells in diapason:
                    for cell in cells:
                        if cell.value == criteria:
                            return cell
            elif hasNumbers(range) == False:
                for cell in diapason:
                    if cell.value == criteria:
                        return cell
        return None

    def getListOfCellsByCriteria(self, criteria, range: str, wsName=""):
        """Finds list of cells with values equal to criteria by
        searching in some range or in whole sheet

        Keyword arguments:
        criteria -- search criteria in cells values
        range -- search range (I22:J22, I, ...), by default set to None
                    so it search in whole sheet
        """
        listOfCells = []

        if hasNumbers(range) == True:
            for cells in self.getWs(wsName)[range]:
                for cell in cells:
                    if cell.value == criteria:
                        listOfCells.append(cell)
        elif hasNumbers(range) == False:
            for cell in self.getWs(wsName)[range]:
                if cell.value == criteria:
                    listOfCells.append(cell)
        return listOfCells

    def unmerge(self, wsName=""):
        for range in self.getWs(wsName).merged_cells.ranges:
            rangeList = list(range.bounds)
            minCol = rangeList[0]
            minRow = rangeList[1]
            maxCol = rangeList[2]
            maxRow = rangeList[3]
            self.getWs(wsName).unmerge_cells(start_row=minRow,
                                    start_column=minCol,
                                    end_row=maxRow,
                                    end_column=maxCol
                                    )
        return

    def merge(self, range: str, wsName=""):
        start = range.split(":")[0]
        end = range.split(":")[1]
        minRow = openpyxl.utils.coordinate_to_tuple(start)[0]
        minCol = openpyxl.utils.coordinate_to_tuple(start)[1]
        maxRow = openpyxl.utils.coordinate_to_tuple(end)[0]
        maxCol = openpyxl.utils.coordinate_to_tuple(end)[1]
        self.getWs(wsName).merge_cells(start_row=minRow,
                            start_column=minCol,
                            end_row=maxRow,
                            end_column=maxCol
                            )
        return

    def mergeByTuple(self, rangeList: list, wsName=""):
        for range in rangeList:
            coord = list(range.bounds)
            rangeStr = str(openpyxl.utils.get_column_letter(coord[0])) + str(
                coord[1]) + ":" + str(openpyxl.utils.get_column_letter(coord[2])) + str(coord[3])
            self.merge(rangeStr)
        return


class Manager:

    def __init__(self):
        self.files = []

    def setWorkDir(self, pathToWorkDir: str):
        self.pathToWorkDir = pathToWorkDir

    def addFileByPath(self, pathToFile: str, fileName: str):
        if ".xlsx" in fileName:
            self.files.append(OpenPyXl(pathToFile, fileName))
        elif ".xls" in fileName:
            self.files.append(PyWin(pathToFile, fileName))

    def addFile(self, file: File):
        self.files.append(file)

    def addFilesInDir(self):
        for r, d, f in os.walk(self.pathToWorkDir):
            for fileName in f:
                self.addFileByPath(self.pathToWorkDir, fileName)
            break

    def __getitem__(self, i: int):
        #print("getitem:", i, len(self.files))
        if i < len(self.files) and i > -1:
            return self.files[i]
        else:
            raise AttributeError

    def removeFile(self, thatFile: File):
        try:
            #print("removeFile:", str(thatFile))
            # print(self.files)
            thatFile.close()
            self.files.remove(thatFile)
        except ValueError:
            print("Couldn`t remove file " + thatFile.fileName)

    def deleteFile(self, thatFile: File, extension=".xlsx"):
        self.removeFile(thatFile)
        fullName = thatFile.pathToFile + "\\"
        fullName += thatFile.fileNameWithoutExtension
        fullName += thatFile.fileExtension
        os.remove(fullName)

    def getNumberOfFiles(self):
        return len(self.files)

    def printAllFiles(self):
        print("\n\n")
        print(self)
        for file in self.files:
            print(file.fileName, file)
        print("\n\n")

    def getFile(self, partOfNameOfFile, extension=".xls"):
        for file in self.files:
            if partOfNameOfFile in file.fileName:
                if extension == file.fileExtension:
                    if file.wasCalled == False:
                        file.wasCalled = True
                        return file
        return None

    def removeUnCalledFiles(self):
        for file in self.files:
            if file.wasCalled == False:
                self.removeFile(file)

    def deleteUnCalledFiles(self):
        for file in self.files:
            if file.wasCalled == False:
                self.deleteFile(file)

    def deleteClosedFiles(self):
        forDelete = []
        for file in self.files:
            if file.isOpened == False:
                forDelete.append(file)

        for file in forDelete:
            self.deleteFile(file)

    def allFromXlsToXlsx(self):
        forRemove = []
        for file in self.files:
            print(file.fileName)
            if file.fileExtension == ".xls":
                saveFileAsXlsx(self, file)
                file.wasCalled = False
                forRemove.append(file)
    
        for file in forRemove:
            self.removeFile(file)


def saveFileAsXlsx(manager: Manager, file: File):
    file.open()
    file.save(file.pathToFile, file.fileNameWithoutExtension, ".xlsx")
    file.close()
    newFileName = file.fileNameWithoutExtension+".xlsx"
    manager.addFileByPath(file.pathToFile, newFileName)



if __name__ == "__main__":

    mng = Manager()
    mng.setWorkDir(r"C:\Users\LuzhanskyiM-Inet\Development\Excel")
    mng.addFilesInDir()

    tmpMng = Manager()
    tmpMng.setWorkDir(r"C:\Users\LuzhanskyiM-Inet\Development\Excel")

    neededFileNames = ["222.xlsx", "111.xls"]
    for neededFileName in neededFileNames:
        for file in mng.files:
            #print("main loop:", file.fileName, str(file))
            #print(file.fileName, neededFileName)
            if file.fileName == neededFileName:
                #print("i`m here")
                #print("if statement:", neededFileName, str(file))
                # mng.removeFile(file)
                tmpMng.addFile(file)

    # mng = tmpMng
    # del tmpMng

    # for file in mng.files:
    #     print(file.fileName)
    #print("privet ebalo")
    # saveFileAsXlsx(mng, mng[1])

    # for file in mng.files:
    #     print(file.fileName)

    # mng[2].open()
    # wb=mng[2].getWb()
    # print(wb["1"].cell(row=2, column=1).value)
