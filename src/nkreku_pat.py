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

        listForTest1 = []
        time1 = time.time()
        rangeIter = "B9" + ":" + "B" + str(VTVWs.max_row)
        for cells in VTVWs[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    company = cell.value

                    columnC = self.findInSaldo(saldoWs, company, None, None, "G")
                    columnD = self.findInSaldo(saldoWs, company, None, None, "H")
                    columnE = self.findInSaldo(saldoWs, company, None, None, "I")
                    
        print(time.time()-time1)

        time1 = time.time()
        rangeIter = "B9" + ":" + "B" + str(VTVWs.max_row)
        for cells in VTVWs[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    company = cell.value

                    column = self.findInSaldo2(saldoWs, company, None, None, "G")
                    
        print(time.time()-time1)



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
                    whatCategory: list, whatResource: list, whatColumn: str, inWhatColumnFind="A"):
        """
        """
        numberOfRows = saldoSheet.max_row
        rangeIter = inWhatColumnFind + "10" + ":" + inWhatColumnFind + str(numberOfRows)
        columnCategory = openpyxl.utils.column_index_from_string("C")
        columnResource = openpyxl.utils.column_index_from_string("F")
        columnWithData = openpyxl.utils.column_index_from_string(whatColumn)
        columnWithData1 = openpyxl.utils.column_index_from_string("G")
        columnWithData2 = openpyxl.utils.column_index_from_string("H")
        columnWithData3 = openpyxl.utils.column_index_from_string("I")

        columnsList = [columnWithData1, columnWithData2, columnWithData3]
        for cells in saldoSheet[rangeIter]:
            for cell in cells:
                # If company name equals to what to find variable
                if cell.value != None and cell.value == whatToFind:
                    # Iterate throug this company data
                    returnValue1 = 0
                    returnValue2 = 0
                    returnValue3 = 0
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
                        value1 = saldoSheet.cell(column=columnsList[0],
                                                row=row).value
                        value2 = saldoSheet.cell(column=columnsList[1],
                                                row=row).value
                        value3 = saldoSheet.cell(column=columnsList[2],
                                                row=row).value
                        if value1 == None:
                            value1 = 0
                        if value2 == None:
                            value2 = 0
                        if value3 == None:
                            value3 = 0
                        # If both category and resource 
                        # are not specified then add all values
                        if not whatCategory and not whatResource:       
                            returnValue1 += value1
                            returnValue2 += value2
                            returnValue3 += value3
                        # If category is specified and resource aren`t
                        elif whatCategory and not whatResource:
                            if category in whatCategory:
                                returnValue1 += value1
                                returnValue2 += value2
                                returnValue3 += value3
                        # If resource is specified and category aren`t
                        elif whatResource and not whatCategory:
                            if resource in whatResource:
                                returnValue1 += value1
                                returnValue2 += value2
                                returnValue3 += value3
                        # If both are specified
                        elif whatCategory and whatResource:
                            if category in whatCategory and resource in whatResource:
                                returnValue1 += value1
                                returnValue2 += value2
                                returnValue3 += value3
                    
                    return returnValue1, returnValue2, returnValue3
        
        # If company wasn`t found in saldo then set return value to 0
        try:
            returnValue1, returnValue2, returnValue3
        except UnboundLocalError:
            returnValue1=0
            returnValue2=0 
            returnValue3 = 0
            return returnValue1, returnValue2, returnValue3
        
        return