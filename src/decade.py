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

        return

    def naselenie(self):
        
        self.decade.open(data_only=False)
        decadeWsNas = self.decade.getWs("Населення")
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()
        self.saldoLastMonth.open(data_only=True)
        saldoLastMontsWs = self.saldoLastMonth.getWs()
        self.gasConsumption.open(data_only=True)
        gasConsumptionWs = self.gasConsumption.getWs("За період")

        self.iterateInTOVandPAT(decadeWsNas, saldoWs, saldoLastMontsWs, gasConsumptionWs)

        return

    def iterateInTOVandPAT(self, decadeSheet, saldoSheet, saldoLastMonthSheet, gasConsumtionSheet):
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
        self.columnFandG(decadeSheet, saldoLastMonthSheet, gasConsumtionSheet, rangeIter)
        self.columnIandK(decadeSheet, rangeIter)
        self.columnJ(decadeSheet, saldoSheet, rangeIter)
        
        # Iterate in column with PAT companies
        rangeIter = "C7" + ":" + "C" + str(numberOfRows)
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        "населення", None, "G", "D")
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        "населення", "2018", "U", "E") 
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        "населення", "2019", "T", "H")
        self.columnFandG(decadeSheet, saldoLastMonthSheet, gasConsumtionSheet, rangeIter)
        self.columnIandK(decadeSheet, rangeIter)
        self.columnJ(decadeSheet, saldoSheet, rangeIter)

        return
        
    
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
        return

    def findInSaldo(self, saldoSheet, whatToFind: str, 
                    whatCategory: str, whatResource: str, whatColumn: str):
        numberOfRows = saldoSheet.max_row
        rangeIter = "A10" + ":" + "A" + str(numberOfRows)
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
        
        # If company wasn`t found in saldo then set return value to 0
        try:
            returnValue
        except UnboundLocalError:
            returnValue = 0
            return returnValue
        
        return

    def columnFandG(self, decadeSheet, saldoSheet, gasConsumptionSheet, rangeIter):
        columnWhereToPutF = openpyxl.utils.column_index_from_string("F")
        columnWhereToPutG = openpyxl.utils.column_index_from_string("G")
        for cells in decadeSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    companyName = cell.value

                    # Working with F column in decade sheet
                    valueFromGasConsumption = self.findInGasConsumption(gasConsumptionSheet, 
                                                                whatToFind=companyName,
                                                                whatColumn="J")
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
                    decadeSheet.cell(column=columnWhereToPutF, row=cell.row).value = summary
                    
                    # Working with G column in decade sheet
                    valueFromGasConsumption = self.findInGasConsumption(gasConsumptionSheet, 
                                                                whatToFind=companyName,
                                                                whatColumn="H")
                    valueFromSaldoLastMonth = self.findInSaldo(saldoSheet,
                                                                whatToFind=companyName,
                                                                whatCategory="населення",
                                                                whatResource=None,
                                                                whatColumn="I")
                    if valueFromGasConsumption == None:
                        valueFromGasConsumption = 0
                    if valueFromSaldoLastMonth == None:
                        valueFromSaldoLastMonth = 0

                    summary = valueFromGasConsumption + valueFromSaldoLastMonth
                    # Where to get price???????????????????????????
                    price = 1
                    decadeSheet.cell(column=columnWhereToPutG, row=cell.row).value = summary * price
        return
        
    
    def findInGasConsumption(self, gasConsumptionSheet, whatToFind: str, whatColumn: str):
        columnWithData = openpyxl.utils.column_index_from_string(whatColumn)
        numberOfRows = gasConsumptionSheet.max_row
        rangeIter = "B13" + ":" + "B" + str(numberOfRows)
        for cells in gasConsumptionSheet[rangeIter]:
            for cell in cells:
                if cell.value != None and cell.value == whatToFind:
                    value = gasConsumptionSheet.cell(column=columnWithData,
                                                    row=cell.row).value
                    if value != None:
                        return value
                    elif value == None:
                        return 0
        return 0

    def columnIandK(self, decadeSheet, rangeIter):
        columnWithPercent = openpyxl.utils.column_index_from_string("I")
        columnWithDebt = openpyxl.utils.column_index_from_string("K")
        columnPaymentForConsumedGas = openpyxl.utils.column_index_from_string("H")
        columnAmountConsumedGas = openpyxl.utils.column_index_from_string("G")
        for cells in decadeSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    paymentForConsumedGas = decadeSheet.cell(column=columnPaymentForConsumedGas, 
                                            row=cell.row).value
                    amountConsumedGas = decadeSheet.cell(column=columnAmountConsumedGas, 
                                            row=cell.row).value
                    
                    # Fill column I in decade
                    if amountConsumedGas == 0:
                        decadeSheet.cell(column=columnWithPercent, 
                                            row=cell.row).value = 0
                    elif amountConsumedGas != 0:
                        decadeSheet.cell(column=columnWithPercent, 
                                            row=cell.row).value = \
                                            paymentForConsumedGas / amountConsumedGas * 100

                    # Fill column K in decade
                    decadeSheet.cell(column=columnWithDebt, 
                                            row=cell.row).value = amountConsumedGas - paymentForConsumedGas
        return

    def columnJ(self, decadeSheet, saldoSheet, rangeIter):

        columnWithDebtPreviousYears = openpyxl.utils.column_index_from_string("D")
        columnConsumedGasAmount = openpyxl.utils.column_index_from_string("G")
        columnPaymentForConsumedGas = openpyxl.utils.column_index_from_string("H")
        columnWhereToPut = openpyxl.utils.column_index_from_string("J")
        for cells in decadeSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    lastPeriodDebt = decadeSheet.cell(column=columnWithDebtPreviousYears,
                                            row=cell.row).value
                    valueFromColumnTSaldo = self.findInSaldo(saldoSheet, 
                                            whatToFind=cell.value, 
                                            whatCategory="населення",
                                            whatResource=None,
                                            whatColumn="T")
                    valueFromColumnTSaldo2019 = self.findInSaldo(saldoSheet, 
                                            whatToFind=cell.value, 
                                            whatCategory="населення",
                                            whatResource="2019",
                                            whatColumn="T")
                    valueFromColumnJSaldo = self.findInSaldo(saldoSheet, 
                                            whatToFind=cell.value, 
                                            whatCategory="населення",
                                            whatResource=None,
                                            whatColumn="J")
                    consumedGasAmount = decadeSheet.cell(column=columnConsumedGasAmount,
                                            row=cell.row).value
                    paymentForConsumedGas = decadeSheet.cell(column=columnPaymentForConsumedGas,
                                            row=cell.row).value
                    
                    total = lastPeriodDebt - valueFromColumnTSaldo + \
                                valueFromColumnTSaldo2019 + valueFromColumnJSaldo + \
                                consumedGasAmount - paymentForConsumedGas
                    
                    decadeSheet.cell(column=columnWhereToPut, row=cell.row).value = total
        return