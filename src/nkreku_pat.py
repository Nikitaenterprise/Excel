from src.alg import *

class NKREKU_PAT(Algorithm):

    def checkIfDirectoryIsReady(self, path: str):
        self.mng.addFilesInDir()

        self.VTV = self.mng.getFile("Шаблон", extension=".xlsx")
        self.mng.getFile("Оборотно-сальдова вiдомiсть")

        self.mng.deleteUnCalledFiles()               
        self.mng.allFromXlsToXlsx()

        try:
            self.saldo = self.mng.getFile("Оборотно-сальдова вiдомiсть",
                                            extension=".xlsx")
            if self.mng.getNumberOfFiles() != self.numberOfFilesToStart:
                raise AttributeError
        except AttributeError:
            print("Не хватает файлов для работы. Проверьте директорию " + str(path))
            msg = r"""Файлы, нужные для работы: 
            1. Оборотно-сальдова вiдомiсть : 
                                за отчетный месяц (1-30 число), 
                                категория ВТВ+НОРМ, 
                                признаки субъекта :
                                            Усі ВАТ+ЗБУТ, 
                                            ВАТ з обов'язками ЗБУТ
                                без лимитов
                                с судами
                    (Менеджер отчетов\Стан розрахунків\
                            Оборотно-сальдовая\
                                Оборотно-сальдова відомість... (2gv))
            Файлы, которые не надо загружать, но они должны быть в папке
            1. Шаблон : пустой шаблон
            После исправления запустите программу заново. 
            Сейчас программа завершит работу
            Нажмите любую клавишу а затем Enter
            """
            print(bcolors.OKGREEN + msg + bcolors.ENDC)
            input()
            exit()

    def deleteFiles(self, programmIsDone=True):
        """Deletes all created files with .xlsx extension
        """
        # If programm has done its work then close files
        if programmIsDone == True:
            try:
                self.saldo.close()
            except:
                print(bcolors.WARNING +\
                    "Программа не смогла закрыть экселевские файлы"\
                    + bcolors.ENDC)
        self.mng.deleteClosedFiles()
        return

    def run(self):
        self.VTV.open(data_only=False)
        VTVWs = self.VTV.getWs()

        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()

        columnToWriteList = ["C", "D", "E"]
        columnList = []
        for column in columnToWriteList:
            columnList.append(openpyxl.utils.column_index_from_string(column))

        
        
        rangeIter = "B9" + ":" + "B" + str(VTVWs.max_row)
        for cells in VTVWs[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    company = cell.value

                    data = self.findInSaldo(saldoWs, company, 
                                            None, None, 
                                            whatColumns=["G", "H", "I"])
                    for i in range(0, len(data)):
                        VTVWs.cell(column=columnList[i], row=cell.row).value = data[i]

                    data = self.findInSaldo(saldoWs, company, 
                                            None, ["2021"], 
                                            whatColumns=["K", "L", "S"])
                    for i in range(0, len(data)):
                        # i + 4 to shift columns from C D E to G H I
                        VTVWs.cell(column=columnList[i]+4, row=cell.row).value = data[i]

                    data = self.findInSaldo(saldoWs, company, 
                                            None, ["!2021"], 
                                            whatColumns=["K", "L", "S"])
                    for i in range(0, len(data)):
                        # i + 4 to shift columns from C D E to K L M
                        VTVWs.cell(column=columnList[i]+8, row=cell.row).value = data[i]

                    data = self.findInSaldo(saldoWs, company, 
                                            None, None, 
                                            whatColumns=["J"])
                    columnP = openpyxl.utils.column_index_from_string("P")
                    VTVWs.cell(column=columnP, row=cell.row).value = data[0]
                    
        self.VTV.save(self.VTV.pathToFile, "Месячная ВТВ+НОРМ", extension=".xlsx")

        self.deleteFiles()

    def findInSaldo(self, saldoSheet, whatToFind: str, 
                    whatCategory: list, whatResource: list, 
                    whatColumns: list, inWhatColumnFind="A"):
        """
        """
        numberOfRows = saldoSheet.max_row
        rangeIter = inWhatColumnFind + "10" + ":" + inWhatColumnFind + str(numberOfRows)
        columnCategory = openpyxl.utils.column_index_from_string("C")
        columnResource = openpyxl.utils.column_index_from_string("F")

        listOfColumns = []
        for column in whatColumns:
            col = openpyxl.utils.column_index_from_string(column)
            listOfColumns.append(col)

        additionListCategory = []
        exclusionListCategory = []
        if whatCategory != None:
            for cat in whatCategory:
                if "!" in cat:
                    exclusionListCategory.append(cat.split("!")[1])
                elif "!" not in cat:
                    additionListCategory.append(cat)
        

        additionListResource = []
        exclusionListResource = []
        if whatResource != None:
            for res in whatResource:
                if "!" in res:
                    exclusionListResource.append(res.split("!")[1]) 
                elif "!" not in res:
                    additionListResource.append(res)
        

        for cells in saldoSheet[rangeIter]:
            for cell in cells:
                # If company name equals to what to find variable
                if cell.value != None and cell.value == whatToFind:
                    # Iterate through this company data
                    row = cell.row

                    returnValuesList = [0]*len(whatColumns)
                    while True:
                        # Move 1 row down
                        row += 1
                        category = saldoSheet.cell(column=columnCategory,
                                                    row=row).value
                        # If its None then we know that company data ends
                        # and other company begins
                        if category == None:
                            break
                        # Transform int value of resource into str 2019 -> "2019"
                        resource = str(saldoSheet.cell(column=columnResource,
                                                        row=row).value).strip()
                        
                        # Get values from different columns
                        # Write them to list
                        valuesList = []
                        for column in listOfColumns:
                            value = saldoSheet.cell(column=column,
                                                    row=row).value
                            if value == None:
                                value = 0
                            valuesList.append(value)
                        
                        willBeCalculatedCategory = False
                        willBeCalculatedResource = False

                        willBeCalculatedCategory = makeDecision(
                                    additionListCategory,
                                    exclusionListCategory,
                                    category)
                        willBeCalculatedResource = makeDecision(
                                    additionListResource,
                                    exclusionListResource,
                                    resource)
                       
                        if (willBeCalculatedCategory and 
                            willBeCalculatedResource):
                            for i in range(0, len(valuesList)):
                                returnValuesList[i] += valuesList[i]
                    
                    return returnValuesList

        # If no data was found then return zeros
        returnValuesList = [0]*len(listOfColumns)
        return returnValuesList


def makeDecision(addList: list, excudeList: list, value):
    
    boolValue = False

    if (not addList and
        not excudeList):

        boolValue = True

    elif (addList and
        not excudeList and
        value in addList):

        boolValue = True

    elif (not addList and
        excudeList and
        value not in excudeList):
        
        boolValue = True

    elif (addList and
        excudeList):

        boolValue = True

    return boolValue