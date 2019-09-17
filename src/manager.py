import os

import openpyxl
import win32com

class File:

    def __init__(self, pathToFile: str, fileName: str):
        self.pathToFile = pathToFile
        self.fileName = fileName
        self.isOpened = False
    
    def open(self):
        pass

    def close(self):
        pass

    def save(self, path: str, name: str, extension=".xlsx"):
        pass

    def isOpen(self):
        return self.isOpened


class PyWin(File):

    def open(self):
        self.excelApp = win32com.client.Dispatch("Excel.Application")
        self.excelApp.Visible = False
        try:
            self.wb = excelApp.Workbooks.Open(self.pathToFile+self.fileName)
            self.isOpen = True
        except:
            self.isOpen = False
            self.excelApp.Quit()
            print("Программа не может открыть файл " + self.fileName)
            raise FileNotFoundError

    def getWb(self):
        if self.isOpen == True:
            return self.wb
    
    def getApp(self):
        if self.isOpen == True:
            return self.excelApp

    def close(self):
        if self.isOpen == True:
            self.wb.Close()
            self.excelApp.Quit()
            self.isOpen = False 

    def save(self, path: str, name: str, extension=".xlsx"):
        if self.isOpen == True:
            if extension == "xlsx":
                # Code for xslx format
                fileFormat = 51 
            elif extension == "xls":
                # Code for xls format
                fileFormat = 56
            self.wb.SaveAs(path+name, FileFormat=fileFormat)


class OpenPyXl(File):
    def open(self, data_only=False, keep_vba=False):
        try:
            self.wb = openpyxl.load_workbook(self.pathToFile+self.fileName, 
                                    data_only=data_only, keep_vba=keep_vba)
            self.isOpen = True
        except:
            self.isOpen = False
            print("Программа не может открыть файл " + self.fileName)
            raise FileNotFoundError

    def getWb(self):
        if self.isOpen == True:
            return self.wb

    def close(self):
        if self.isOpen == True: 
            self.wb.close()
            self.isOpen = False
    
    def save(self, path: str, name: str, extension=".xlsx"):
        if self.isOpen == True:
            self.wb.save(path+name)

class Manager:

    def __init__(self):
        self.files = []

    def setWorkDir(self, pathToWorkDir: str):
        self.pathToWorkDir = pathToWorkDir
        
    def addFileByPath(self, pathToFile: str, fileName: str):
        self.files.append(File(pathToFile, fileName))

    def addFile(self, file: File):
        self.files.append(file)

    def addFilesInDir(self):
        for r, d, f in os.walk(self.pathToWorkDir):
            for fileName in f:
                if ".xlsx" in fileName:
                    self.files.append(OpenPyXl(self.pathToWorkDir, fileName))
                elif ".xls" in fileName: 
                    self.files.append(PyWin(self.pathToWorkDir, fileName))
            break

    def __getitem__(self, i: int):
        print("getitem:", i, len(self.files))
        if i < len(self.files):
            return self.files[i]
        else:
            raise AttributeError

    def removeFile(self, thatFile: File):
        #try:
        print("removeFile:", str(thatFile))
        #print(self.files)
        self.files.remove(thatFile)
        thatFile.close()
        #except ValueError:
            #print("Couldn`t remove file " + thatFile.fileName)

    def getNumberOfFiles(self):
        return len(self.files)

def iterManager(manager):
    for i in range(0, manager.getNumberOfFiles()):
        #print("iterManager: " + str(manager[i]))
        yield manager[i]
        #i += 1

def openFile(file: File):
    #extension = os.path.splitext(file.fileName)[1] # may be .xls or .xlsx
    extension = file.fileName.split(".")[1]
    extension = "." + extension

    if extension == ".xls":
        # excelApp = win32com.client.Dispatch("Excel.Application")
        # excelApp.Visible = False

        # try:
        #     wb = excelApp.Workbooks.Open(os.path.abspath(name))
        # except:
        #     excelApp.Quit()
        #     print("Программа не может открыть файл " + name)
        #     raise FileNotFoundError
            
        self.saveAsXlsx(newName, wb, excelApp)

        self.fileNameWithPath = newName
            
        extension = os.path.splitext(self.fileNameWithPath)[1] # reinitialize extension
    if extension == ".xlsx":
        self.wb = openpyxl.load_workbook(self.fileNameWithPath, data_only=self.data_only, keep_vba=self.keep_vba)
        self.ws = self.wb[self.wb.sheetnames[self.worksheetNumberInBook]]
    elif how == "pywin":
        self.excelApp = win32com.client.Dispatch("Excel.Application")
        self.excelApp.Visible = False
        try:
            self.wbPW = excelApp.Workbooks.Open(os.path.abspath(self.fileNameWithPath))
        except:
            self.excelApp.Quit()
            print("Программа не может открыть файл " + self.fileNameWithPath)
            raise FileNotFoundError

mng = Manager()
mng.setWorkDir(r"C:\Users\LuzhanskyiM-Inet\Development\Excel")
mng.addFilesInDir()

tmpMng = Manager()
tmpMng.setWorkDir(r"C:\Users\LuzhanskyiM-Inet\Development\Excel")

neededFileNames = ["111.xlsx", "222.xlsx"]
for neededFileName in neededFileNames:
    for file in iterManager(mng):
        #print("main loop:", file.fileName, str(file))
        print(file.fileName, neededFileName)
        if file.fileName == neededFileName:
            print("i`m here")
            #print("if statement:", neededFileName, str(file))
            #mng.removeFile(file)
            tmpMng.addFile(file)

for file in iterManager(tmpMng):
    print(file.fileName)

mng = tmpMng
del tmpMng
print("\n\n\n")
for file in mng.files:
    print(file.fileName, mng.getNumberOfFiles())
    mng.removeFile(file)

for file in iterManager(mng):
    print(file.fileName)



    