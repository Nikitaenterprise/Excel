from src.alg import *

class Decade(Algorithm):

    def checkIfDirectoryIsReady(self, path: str):
        self.mng.addFilesInDir()

        # Checks if the files are present
        self.decade = self.mng.getFile("Декадка", extension=".xlsx")
        self.decade.shouldBeDeleted = False
        self.mng.getFile("Оборотно-сальдова вiдомiсть", exactMatch=True)
        self.mng.getFile("Оборотно-сальдова вiдомiсть последний месяц", exactMatch=True)
        self.mng.getFile("gpg")

        # Deletes other files in dir
        self.mng.deleteUnCalledFiles()               
        self.mng.allFromXlsToXlsx()

        try:
            
            self.saldo = self.mng.getFile("Оборотно-сальдова вiдомiсть", 
                                    extension=".xlsx", exactMatch=True)
            self.saldoLastMonth = self.mng.getFile("Оборотно-сальдова вiдомiсть последний месяц", 
                                    extension=".xlsx", exactMatch=True)
            self.gasConsumption = self.mng.getFile("gpg", extension=".xlsx")
            
            if self.mng.getNumberOfFiles() != self.numberOfFilesToStart:
                    raise AttributeError
        except AttributeError:
            print("Не хватает файлов для работы. Проверьте директорию " + str(path))
            msg = """11111
            """
            print(bcolors.OKGREEN + msg + bcolors.ENDC)
            input()
            exit()
            
    def deleteFiles(self, programmIsDone=True):
        """Deletes all created files with .xlsx extension
        """
        # If programm has daone its work then close files
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
        self.naselenie()
        self.decade.save(self.decade.pathToFile, "111", extension=".xlsx")
        self.deleteFiles()

    def naselenie(self):
        
        self.decade.open(data_only=False)
        decadeWsNas = self.decade.getWs("Населення")
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()
        self.saldoLastMonth.open(data_only=True)
        saldoLastMontsWs = self.saldoLastMonth.getWs()
        self.gasConsumption = open(data_only=True)
        gasConsumptionWs = self.gasConsumption.getWs("За період")

        self.iterateInTOVandPAT(decadeWsNas, saldoWs)

    def iterateInTOVandPAT(self, decadeSheet, saldoSheet):
        """
        """
        numberOfRows = decadeSheet.max_row
        # Iterate in column with TOV companies
        rangeIter = "B7" + ":" + "B" + str(numberOfRows)
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        "населення", None, "G", "D")
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        "населення", "2018", "U", "E")
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        "населення", "2019", "T", "H")
        
        # Iterate in column with PAT companies
        rangeIter = "C7" + ":" + "C" + str(numberOfRows)
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        "населення", None, "G", "D")
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        "населення", "2018", "U", "E") 
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        "населення", "2019", "T", "H")
        
    
    def findInSaldoWriteToDecade(self, decadeSheet, saldoSheet, rangeIter, 
                                    whatCategory: str, whatResource: str, 
                                    whatColumn: str, whereToPut: str):
        """
        """
        columnWhereToPut = openpyxl.utils.column_index_from_string(whereToPut)
        for cells in decadeSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    companyName = cell.value
                    value = self.findInSaldo(saldoSheet, 
                                            whatToFind=companyName,
                                            whatCategory=whatCategory,
                                            whatResource=whatResource,
                                            whatColumn=whatColumn)
                    if value != None:
                        decadeSheet.cell(column=columnWhereToPut, 
                                        row=cell.row).value = value

    def debtForPreviousPeriod(self):
        pass

    def findInSaldo(self, saldoSheet, whatToFind: str, 
                    whatCategory: str, whatResource: str, whatColumn: str):
        numberOfRows = saldoSheet.max_row
        rangeIter = "A10" + ":" + "A" + str(numberOfRows)
        columnCategory = openpyxl.utils.column_index_from_string("C")
        columnResource = openpyxl.utils.column_index_from_string("F")
        columnWithData = openpyxl.utils.column_index_from_string(whatColumn)
        for cells in saldoSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    # If company name equals to what to find variable
                    if cell.value == whatToFind:
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

                            if whatCategory != None and category == whatCategory:
                                # If resource is not specified
                                if whatResource == None:
                                    value = saldoSheet.cell(column=columnWithData,
                                                    row=row).value
                                    if value != None:
                                        returnValue += value

                                elif whatResource != None:
                                    # Get resource value and make it integer
                                    resource = saldoSheet.cell(column=columnResource,
                                                            row=row).value

                                    if int(resource) == int(whatResource):
                                        value = saldoSheet.cell(column=columnWithData,
                                                        row=row).value
                                        if value != None:
                                            returnValue += value
                                    elif int(resource) != int(whatResource):
                                        continue

                            elif category != whatCategory:
                                continue
                        
                        return returnValue

    def columnFandG(self, decadeSheet, saldoSheet, rangeIter):
        columnWhereToPut = openpyxl.utils.column_index_from_string("F")
        for cells in decadeSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    companyName = cell.value
                    valueFromGasConsumption = self.findInGasConsumption(whatToFind=companyName)
                    valueFromSaldoLastMonth = self.findInSaldo(saldoSheet,
                                                                whatToFind=companyName,
                                                                whatCategory="населення",
                                                                whatResource=None,
                                                                whatColumn="H")
                    if valueFromGasConsumption == None:
                        valueFromGasConsumption = 0
                    if valueFromSaldoLastMonth == None:
                        valueFromSaldoLastMonth = 0
                    summary = valueFromGasConsumption + valueFromSaldoLastMonth
                    decadeSheet.cell(column=columnWhereToPut, row=cell.row).value = summary
        pass
    
    def findInGasConsumption(self, whatToFind: str):

        pass