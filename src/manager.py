import os

import openpyxl
import win32com.client


class File:

    def __init__(self, pathToFile: str, fileName: str):
        self.pathToFile = pathToFile
        self.fileName = fileName
        self.fileExtension = os.path.splitext(self.pathToFile + "\\" + self.fileName)[1]
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
            self.wb.SaveAs(path+name, FileFormat=fileFormat)


class OpenPyXl(File):
    def open(self, data_only=False, keep_vba=False):
        if self.isOpened == False:
            try:
                self.wb = openpyxl.load_workbook(self.pathToFile+self.fileName,
                                                 data_only=data_only, keep_vba=keep_vba)
                self.isOpened = True
            except:
                self.isOpened = False
                print("Программа не может открыть файл " + self.fileName)
                raise FileNotFoundError

    def getWb(self):
        if self.isOpened == True:
            return self.wb

    def close(self):
        if self.isOpened == True:
            self.wb.close()
            self.isOpened = False

    def save(self, path: str, name: str, extension=".xlsx"):
        if self.isOpened == True:
            self.wb.save(path+name)


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
            self.files.remove(thatFile)
            thatFile.close()
        except ValueError:
            print("Couldn`t remove file " + thatFile.fileName)

    def deleteFile(self, thatFile: File, extension=".xlsx"):
        self.removeFile(thatFile)
        fullName = thatFile.pathToFile + "\\"
        fullName += thatFile.fileName.split(".")[0] + "."
        fullName += extension.split(".")[1]
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

def saveFileAsXlsx(manager: Manager, file: File):
    newFileName = file.fileName.split(".")[0]
    file.open()
    file.save(file.pathToFile, newFileName, ".xlsx")
    file.close()
    manager.addFileByPath(file.pathToFile, str(newFileName)+".xlsx")


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
