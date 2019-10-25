from src.alg import *

class NKREKU_PAT(Algorithm):

    def checkIfDirectoryIsReady(self, path: str):
        self.mng.addFilesInDir()

        self.VTV = self.mng.getFile("ВТВ", extension=".xlsx")
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
                                            признаки субъекта Усі ВАТ+ЗБУТ, ВАТ з обов'язками ЗБУТ
                        (Менеджер отчетов\Стан розрахунків\Оборотно-сальдовая\Оборотно-сальдова відомість... (2gv))
            Файлы, которые не надо загружать, но они должны быть в папке
            1. ВТВ : пустой шаблон
            После исправления запустите программу заново. Сейчас программа завершит работу
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

                    data = self.findInSaldo2(saldoWs, company, None, None, whatColumns=["G", "H", "I"])
                    for i in range(0, len(data)):
                        VTVWs.cell(column=columnList[i], row=cell.row).value = data[i]

                    data = self.findInSaldo2(saldoWs, company, None, ["2019"], whatColumns=["K", "L", "S"])
                    for i in range(0, len(data)):
                        # i + 4 to shift columns from C D E to G H I
                        VTVWs.cell(column=columnList[i]+4, row=cell.row).value = data[i]

                    data = self.findInSaldo2(saldoWs, company, None, ["!2019"], whatColumns=["K", "L", "S"])
                    for i in range(0, len(data)):
                        # i + 4 to shift columns from C D E to K L M
                        VTVWs.cell(column=columnList[i]+8, row=cell.row).value = data[i]
                    
        self.VTV.save(self.VTV.pathToFile, "111", extension=".xlsx")


    def findInSaldo(self, saldoSheet, whatToFind: str, 
                    whatCategory: list, whatResource: list, whatColumn: str, inWhatColumnFind="A"):
        """
        """
        numberOfRows = saldoSheet.max_row
        rangeIter = inWhatColumnFind + "10" + ":" + inWhatColumnFind + str(numberOfRows)
        columnCategory = openpyxl.utils.column_index_from_string("C")
        columnResource = openpyxl.utils.column_index_from_string("F")
        columnWithData = openpyxl.utils.column_index_from_string(whatColumn)
        for cells in saldoSheet[rangeIter]:
            for cell in cells:
                # If company name equals to what to find variable
                if cell.value != None and cell.value == whatToFind:
                    # Iterate throug this company data
                    returnValue = 0
                    row = cell.row
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
                        value = saldoSheet.cell(column=columnWithData,
                                                row=row).value
                        if value != None:
                            # If both category and resource 
                            # are not specified then add all values
                            if not whatCategory and not whatResource:       
                                returnValue += value
                            # If category is specified and resource aren`t
                            elif whatCategory and not whatResource:
                                if category in whatCategory:
                                    returnValue += value
                            # If resource is specified and category aren`t
                            elif whatResource and not whatCategory:
                                if resource in whatResource:
                                    returnValue += value
                            # If both are specified
                            elif whatCategory and whatResource:
                                if category in whatCategory and resource in whatResource:
                                    returnValue += value
                    
                    return returnValue
        
        # If company wasn`t found in saldo then set return value to 0
        try:
            returnValue
        except UnboundLocalError:
            returnValue = 0
            return returnValue
        
        return

    def findInSaldo2(self, saldoSheet, whatToFind: str, 
                    whatCategory: list, whatResource: list, whatColumns: list, inWhatColumnFind="A"):
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

        exclusionListCategory = []
        if whatCategory != None:
            for cat in whatCategory:
                if "!" in cat:
                    catList = cat.split("!")
                    exclusionListCategory.append(catList[1])
        if not exclusionListCategory:
            exclusionListCategory.append("")

        exclusionListResource = []
        if whatResource != None:
            for res in whatResource:
                if "!" in res:
                    resList = res.split("!")
                    exclusionListResource.append(resList[1]) 
        if not exclusionListResource:
            exclusionListResource.append("")

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
                        
                        willBeCalculated = False

                        # If just category or resource was specified without excluding
                        # values. for example ["2019"] or ["ТЕ"]
                        #if "!" not in whatCategory and "!" not in whatResource:
                        # If both category and resource 
                        # are not specified then add all values
                        if not whatCategory and not whatResource:       
                            willBeCalculated = True
                        # If category is specified and resource aren`t
                        elif whatCategory and not whatResource:
                            if category in whatCategory or category not in exclusionListCategory:
                                willBeCalculated = True
                            elif category in exclusionListCategory:
                                willBeCalculated = False

                        # If resource is specified and category aren`t
                        elif whatResource and not whatCategory:
                            if resource in whatResource or resource not in exclusionListResource:
                                willBeCalculated = True
                            if resource in exclusionListResource:
                                willBeCalculated = False
                        # If both are specified
                        elif whatCategory and whatResource:
                            if category in whatCategory and\
                                    resource in whatResource or\
                                    category not in exclusionListCategory and\
                                    resource not in exclusionListResource:
                                willBeCalculated = True
                            if category in exclusionListCategory or\
                                    resource in exclusionListResource:
                                willBeCalculated = False
                            

                        if willBeCalculated == True:
                            for i in range(0, len(valuesList)):
                                returnValuesList[i] += valuesList[i]
                        # If excluding values are in resource or category
                        # For example ["!2019"] or ["!ТЕ"]
                        #elif "!" in whatCategory or "!" in whatResource:

                    
                    return returnValuesList

        # If no data waas found then return zeros
        returnValuesList = [0]*len(listOfColumns)
        return returnValuesList
