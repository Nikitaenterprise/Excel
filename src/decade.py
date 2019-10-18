from src.alg import *

class Decade(Algorithm):

    def checkIfDirectoryIsReady(self, path: str):
        self.mng.addFilesInDir()

        # Checks if the files are present
        self.decade = self.mng.getFile("Декадка", 
                                        extension=".xlsx")
        self.decade.shouldBeDeleted = False
        self.tkePrev = self.mng.getFile("ТКЕ начало года", 
                                        exactMatch=True, extension=".xlsx")
        self.tkePrev.shouldBeDeleted = False
        self.tkeDK = self.mng.getFile("ТКЕ ДК", 
                                        extension=".xlsx", exactMatch=True)
        self.tkeDK.shouldBeDeleted = False
        self.promDK = self.mng.getFile("Пром ДК", 
                                        extension=".xlsx", exactMatch=True)
        self.promDK.shouldBeDeleted = False
        self.promPrev = self.mng.getFile("Розрахунок промислових минулі роки", 
                                        extension=".xlsx", exactMatch=True)
        self.promPrev.shouldBeDeleted = False
                                        
        
        self.mng.getFile("Оборотно-сальдова вiдомiсть", exactMatch=True)
        self.mng.getFile("Оборотно-сальдова вiдомiсть последний месяц", exactMatch=True)
        self.mng.getFile("ТКЕ", exactMatch=True)
        self.mng.getFile("gpg")
        self.mng.getFile("Промисловість_")

        # Deletes other files in dir
        self.mng.deleteUnCalledFiles()               
        self.mng.allFromXlsToXlsx()

        try:
            
            self.saldo = self.mng.getFile("Оборотно-сальдова вiдомiсть", 
                                    extension=".xlsx", exactMatch=True)
            self.saldoLastMonth = self.mng.getFile("Оборотно-сальдова вiдомiсть последний месяц", 
                                    extension=".xlsx", exactMatch=True)
            self.tke = self.mng.getFile("ТКЕ", 
                                    extension=".xlsx", exactMatch=True)
            self.gasConsumption = self.mng.getFile("gpg", 
                                    extension=".xlsx")
            self.prom = self.mng.getFile("Промисловість_", 
                                    extension=".xlsx")
            
            
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
        # self.naselenie()
        # self.religion()
        # self.budget()
        # self.teploseti()
        self.promishlennost()
        self.decade.save(self.decade.pathToFile, "111", extension=".xlsx")
        self.deleteFiles()

        return

    def naselenie(self):
        
        self.decade.open(data_only=False)
        decadeWsNas = self.decade.getWs("Населення")
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()
        self.saldoLastMonth.open(data_only=True)
        saldoLastMonthWs = self.saldoLastMonth.getWs()
        self.gasConsumption.open(data_only=True)
        gasConsumptionWs = self.gasConsumption.getWs("За період")

        self.naselenieIterateInTOVandPAT(decadeWsNas, saldoWs, saldoLastMonthWs, gasConsumptionWs)

        return

    def religion(self):

        self.decade.open(data_only=False)
        decadeWsReligion = self.decade.getWs("Релігія")
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()
        self.saldoLastMonth.open(data_only=True)
        saldoLastMonthWs = self.saldoLastMonth.getWs()
        self.gasConsumption.open(data_only=True)
        gasConsumptionWs = self.gasConsumption.getWs("За період")

        self.religionIterateInTOVandPAT(decadeWsReligion, saldoWs, saldoLastMonthWs, gasConsumptionWs)

        return

    def budget(self):

        self.decade.open(data_only=False)
        decadeWsBudget = self.decade.getWs("Бюджет")
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()
        self.saldoLastMonth.open(data_only=True)
        saldoLastMonthWs = self.saldoLastMonth.getWs()
        self.gasConsumption.open(data_only=True)
        gasConsumptionWs = self.gasConsumption.getWs("За період")

        self.budgetIterateInTOVandPAT(decadeWsBudget, saldoWs, saldoLastMonthWs, gasConsumptionWs)

        return

    def teploseti(self):

        self.decade.open(data_only=False)
        decadeWsTeploseti = self.decade.getWs("Тепломережі")
        self.tkeDK.open(data_only=True)
        tkeDKWs = self.tkeDK.getWs()
        self.tkePrev.open(data_only=True)
        tkePrevWs = self.tkePrev.getWs()
        self.tke.open(data_only=True)
        tkeWs = self.tke.getWs("База")

        rangeIterInDecade = "A9" + ":" + "A" + str(decadeWsTeploseti.max_row)
        self.teplosetiIterInRegions(decadeWsTeploseti, tkePrevWs, tkeWs, tkeDKWs, rangeIterInDecade)

    def promishlennost(self):
        
        self.decade.open(data_only=False)
        decadeWsProm = self.decade.getWs("Промисловість")
        self.promDK.open(data_only=True)
        promDKWs = self.promDK.getWs()
        self.promPrev.open(data_only=True)
        promPrevWs = self.promPrev.getWs()
        self.prom.open(data_only=True)
        promWs = self.prom.getWs("База_2")
        # Delete Naftogaz trading data
        self.deleteCompanyData(promWs, ["42399676"])
        self.prom.save(self.prom.pathToFile, "qqqqq", extension=".xlsx")
        #rangeIterInDecade = "A9" + ":" + "A" + str(decadeWsProm.max_row)
        #self.promIterInRegions(decadeWsProm, promDKWs, promWs, promPrevWs, rangeIterInDecade)

    def promIterInRegions(self, decadeSheet, promDkSheet, promSheet, promPrevSheet, rangeIter):
        """
        """
        for cells in decadeSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    region = cell.value

                    listOfTwoValues = self.promColumnB(promPrevSheet, promDkSheet, region)
                    # Get debt of NAK and DK together
                    summOfTwoValues = listOfTwoValues[0] + listOfTwoValues[1]
                    value1 = listOfTwoValues[1]
                    decadeSheet.cell(column=cell.column+1, row=cell.row).value = summOfTwoValues

                    value2 = self.promColumnsFromCToF(promSheet, region, "DP")
                    decadeSheet.cell(column=cell.column+2, row=cell.row).value = value2
                    
                    decadeSheet.cell(column=cell.column+3, row=cell.row).value =\
                                    self.promColumnsFromCToF(promSheet, region, "CX")
                    
                    value3 = self.promColumnsFromCToF(promSheet, region, "CY")
                    decadeSheet.cell(column=cell.column+4, row=cell.row).value = value3
                    
                    value4 = self.promColumnsFromCToF(promSheet, region, "CZ")
                    decadeSheet.cell(column=cell.column+5, row=cell.row).value = value4

                    if value3 != 0:
                        decadeSheet.cell(column=cell.column+6, row=cell.row).value = value4 / value3 * 100
                    elif value3 == 0:
                        decadeSheet.cell(column=cell.column+6, row=cell.row).value = 0

                    value5 = self.promColumnsFromCToF(promSheet, region, "AB")
                    value6 = self.promColumnsFromCToF(promSheet, region, "DO")
                    total = value5 + value6 + value2 + value1 + value3 + value4
                    decadeSheet.cell(column=cell.column+7, row=cell.row).value = total
                    decadeSheet.cell(column=cell.column+8, row=cell.row).value = value3 - value4

    
    def naselenieIterateInTOVandPAT(self, decadeSheet, saldoSheet, saldoLastMonthSheet, gasConsumtionSheet):
        """
        """
        numberOfRows = decadeSheet.max_row
        # Iterate in column with TOV companies
        rangeIter = "B7" + ":" + "B" + str(numberOfRows)
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["населення"], None, "G", "D")
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["населення"], ["2018"], "U", "E")
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["населення"], ["2019"], "T", "H")
        self.columnF(decadeSheet, saldoLastMonthSheet, gasConsumtionSheet, 
                    rangeIter, "J", ["населення"])
        self.columnG(decadeSheet, saldoLastMonthSheet, gasConsumtionSheet,
                    rangeIter, "H", ["населення"])
        self.columnIandK(decadeSheet, rangeIter)
        self.columnJ(decadeSheet, saldoSheet, rangeIter, ["населення"])
        
        # Iterate in column with PAT companies
        rangeIter = "C7" + ":" + "C" + str(numberOfRows)
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["населення"], None, "G", "D")
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["населення"], ["2018"], "U", "E") 
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["населення"], ["2019"], "T", "H")
        self.columnF(decadeSheet, saldoLastMonthSheet, gasConsumtionSheet, 
                    rangeIter, "J", ["населення"])
        self.columnG(decadeSheet, saldoLastMonthSheet, gasConsumtionSheet,
                    rangeIter, "H", ["населення"])
        self.columnIandK(decadeSheet, rangeIter)
        self.columnJ(decadeSheet, saldoSheet, rangeIter, ["населення"])

        return

    def religionIterateInTOVandPAT(self, decadeSheet, saldoSheet, saldoLastMonthSheet, gasConsumtionSheet):
        """
        """
        numberOfRows = decadeSheet.max_row
        # Iterate in column with TOV companies
        rangeIter = "B7" + ":" + "B" + str(numberOfRows)
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["релігійні організації", "вічний вогонь"], None, "G", "D")
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["релігійні організації", "вічний вогонь"], ["2018"], "U", "E")
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["релігійні організації", "вічний вогонь"], ["2019"], "T", "H")
        self.columnF(decadeSheet, saldoLastMonthSheet, gasConsumtionSheet, 
                    rangeIter, "W", ["релігійні організації", "вічний вогонь"])
        self.columnG(decadeSheet, saldoLastMonthSheet, gasConsumtionSheet,
                    rangeIter, "W", ["релігійні організації", "вічний вогонь"])
        self.columnIandK(decadeSheet, rangeIter)
        self.columnJ(decadeSheet, saldoSheet, rangeIter, ["релігійні організації", "вічний вогонь"])
        
        # Iterate in column with PAT companies
        rangeIter = "C7" + ":" + "C" + str(numberOfRows)
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["релігійні організації", "вічний вогонь"], None, "G", "D")
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["релігійні організації", "вічний вогонь"], ["2018"], "U", "E") 
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["релігійні організації", "вічний вогонь"], ["2019"], "T", "H")
        self.columnF(decadeSheet, saldoLastMonthSheet, gasConsumtionSheet, 
                    rangeIter, "W", ["релігійні організації", "вічний вогонь"])
        self.columnG(decadeSheet, saldoLastMonthSheet, gasConsumtionSheet,
                    rangeIter, "W", ["релігійні організації", "вічний вогонь"])
        self.columnIandK(decadeSheet, rangeIter)
        self.columnJ(decadeSheet, saldoSheet, rangeIter, ["релігійні організації", "вічний вогонь"])

        return

    def budgetIterateInTOVandPAT(self, decadeSheet, saldoSheet, saldoLastMonthSheet, gasConsumtionSheet):
        """
        """
        numberOfRows = decadeSheet.max_row
        # Iterate in column with TOV companies
        rangeIter = "B7" + ":" + "B" + str(numberOfRows)
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["бюджет"], None, "G", "D")
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["бюджет"], ["2018"], "U", "E")
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["бюджет"], ["2019"], "T", "H")
        self.columnF(decadeSheet, saldoLastMonthSheet, gasConsumtionSheet, 
                    rangeIter, None, ["бюджет"])
        self.columnG(decadeSheet, saldoLastMonthSheet, gasConsumtionSheet,
                    rangeIter, None, ["бюджет"])
        self.columnIandK(decadeSheet, rangeIter)
        self.columnJ(decadeSheet, saldoSheet, rangeIter, ["бюджет"])
        
        # Iterate in column with PAT companies
        rangeIter = "C7" + ":" + "C" + str(numberOfRows)
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["бюджет"], None, "G", "D")
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["бюджет"], ["2018"], "U", "E") 
        self.findInSaldoWriteToDecade(decadeSheet, saldoSheet, rangeIter,
                                        ["бюджет"], ["2019"], "T", "H")
        self.columnF(decadeSheet, saldoLastMonthSheet, gasConsumtionSheet, 
                    rangeIter, None, ["бюджет"])
        self.columnG(decadeSheet, saldoLastMonthSheet, gasConsumtionSheet,
                    rangeIter, None, ["бюджет"])
        self.columnIandK(decadeSheet, rangeIter)
        self.columnJ(decadeSheet, saldoSheet, rangeIter, ["бюджет"])

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
                    whatCategory: list, whatResource: list, whatColumn: str):
        """
        """
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
                        # Transform int value of resource into str 2019 -> "2019"
                        resource = str(saldoSheet.cell(column=columnResource,
                                                row=row).value)
                        value = saldoSheet.cell(column=columnWithData,
                                                row=row).value
                        if value != None:
                            # If both category and resource 
                            # are not specified then add all values
                            if not whatCategory and not whatResource:       
                                returnValue += value
                            # If category is specified and resource aren`t
                            if whatCategory and not whatResource:
                                if category in whatCategory:
                                    returnValue += value
                            # If resource is specified and category aren`t
                            if whatResource and not whatCategory:
                                if resource in whatResource:
                                    returnValue += value
                            # If both are specified
                            if whatCategory and whatResource:
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
        
    def columnF(self, decadeSheet, saldoSheet, gasConsumptionSheet, 
                rangeIter: str, whatColumnInConsumption: str, whatCategory: list):
        """
        """
        columnWhereToPut = openpyxl.utils.column_index_from_string("F")
        for cells in decadeSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    companyName = cell.value

                    # Working with F column in decade sheet
                    if whatColumnInConsumption != None:
                        valueFromGasConsumption = self.findInGasConsumption(gasConsumptionSheet, 
                                                                whatToFind=companyName,
                                                                whatColumn=whatColumnInConsumption)
                    elif whatColumnInConsumption == None:
                        valueFromGasConsumption = 0

                    valueFromSaldoLastMonth = self.findInSaldo(saldoSheet,
                                                                whatToFind=companyName,
                                                                whatCategory=whatCategory,
                                                                whatResource=None,
                                                                whatColumn="H")
                    if valueFromGasConsumption == None:
                        valueFromGasConsumption = 0
                    if valueFromSaldoLastMonth == None:
                        valueFromSaldoLastMonth = 0

                    summary = valueFromGasConsumption + valueFromSaldoLastMonth
                    decadeSheet.cell(column=columnWhereToPut, row=cell.row).value = summary
        return

    def columnG(self, decadeSheet, saldoSheet, gasConsumptionSheet,
                rangeIter: str, whatColumnInConsumption: str, whatCategory: list):
        """
        """
        columnWhereToPut = openpyxl.utils.column_index_from_string("G")
        for cells in decadeSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    companyName = cell.value

                    # Working with G column in decade sheet
                    if whatColumnInConsumption != None:
                        valueFromGasConsumption = self.findInGasConsumption(gasConsumptionSheet, 
                                                                whatToFind=companyName,
                                                                whatColumn=whatColumnInConsumption)
                    elif whatColumnInConsumption == None:
                        valueFromGasConsumption = 0

                    valueFromSaldoLastMonth = self.findInSaldo(saldoSheet,
                                                                whatToFind=companyName,
                                                                whatCategory=whatCategory,
                                                                whatResource=None,
                                                                whatColumn="I")
                    if valueFromGasConsumption == None:
                        valueFromGasConsumption = 0
                    if valueFromSaldoLastMonth == None:
                        valueFromSaldoLastMonth = 0

                    # Where to get price???????????????????????????
                    price = 1
                    summary = valueFromGasConsumption + valueFromSaldoLastMonth
                    decadeSheet.cell(column=columnWhereToPut, row=cell.row).value = summary * price
        return


    def findInGasConsumption(self, gasConsumptionSheet, whatToFind: str, whatColumn: str):
        """
        """
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
        """
        """
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

    def columnJ(self, decadeSheet, saldoSheet, rangeIter: str, whatCategory: list):
        """
        """
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
                                            whatCategory=whatCategory,
                                            whatResource=None,
                                            whatColumn="T")
                    valueFromColumnTSaldo2019 = self.findInSaldo(saldoSheet, 
                                            whatToFind=cell.value, 
                                            whatCategory=whatCategory,
                                            whatResource=["2019"],
                                            whatColumn="T")
                    valueFromColumnJSaldo = self.findInSaldo(saldoSheet, 
                                            whatToFind=cell.value, 
                                            whatCategory=whatCategory,
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

    def teplosetiIterInRegions(self, decadeSheet, tkePrevWs, tkeWs, tkeDKsheet, rangeIter):
        """
        """
        for cells in decadeSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    region = cell.value

                    listOfTwoValues = self.teplosetiColumnB(tkePrevWs, tkeDKsheet, region)
                    # Get debt of NAK and tkeDK together
                    summOfTwoValues = listOfTwoValues[0] + listOfTwoValues[1]
                    decadeSheet.cell(column=cell.column+1, row=cell.row).value = summOfTwoValues
                                    
                    
                    value1 = self.teplosetiColumnsFromCToF(tkeWs, region, "DI")
                    decadeSheet.cell(column=cell.column+2, row=cell.row).value = value1
                                    
                    decadeSheet.cell(column=cell.column+3, row=cell.row).value = \
                                    self.teplosetiColumnsFromCToF(tkeWs, region, "DJ")
                    
                    value2 = self.teplosetiColumnsFromCToF(tkeWs, region, "DK")
                    decadeSheet.cell(column=cell.column+4, row=cell.row).value = value2

                    value3 = self.teplosetiColumnsFromCToF(tkeWs, region, "DQ")
                    decadeSheet.cell(column=cell.column+5, row=cell.row).value = value3
                                    
                    value4 = self.teplosetiColumnsFromCToF(tkeWs, region, "CU")
                    value5 = self.teplosetiColumnsFromCToF(tkeWs, region, "CG")
                    value6 = listOfTwoValues[1]
                    total = value1 + value4 + value5 + value6 + value2 - value3

                    decadeSheet.cell(column=cell.column+7, row=cell.row).value = total
                    
                    if value2 != 0:
                        decadeSheet.cell(column=cell.column+6, row=cell.row).value = value3 / value2 * 100
                    elif value2 == 0:
                        decadeSheet.cell(column=cell.column+6, row=cell.row).value = 0

                    decadeSheet.cell(column=cell.column+8, row=cell.row).value = value2 - value3


                    
    def teplosetiColumnB(self, tkeSheet, tkeDKsheet, regionName):
        """
        """
        rangeIterInTkePrev = "B15" + ":" + "B" + str(tkeSheet.max_row)
        columnDebtInTke = openpyxl.utils.column_index_from_string("EO")
        returnValueFromTke = 0
        isDone = False
        for cells in tkeSheet[rangeIterInTkePrev]:
            for cell in cells:
                # If left cell is equal to ВАТ/ПАТ
                if tkeSheet.cell(column=cell.column-1, row=cell.row).value == "ВАТ/ПАТ":
                    isDone = True
                    break
                
                if cell.value == regionName:
                    returnValueFromTke = tkeSheet.cell(column=columnDebtInTke, row=cell.row).value
                    isDone = True
                    break
            if isDone:
                break

        rangeIterInTkeDK = "B12" + ":" + "B" + str(tkeDKsheet.max_row)
        columnDebtInTkeDK = openpyxl.utils.column_index_from_string("EG")
        returnValueFromTkeDK = 0
        isDone = False
        for cells in tkeDKsheet[rangeIterInTkeDK]:
            for cell in cells:
                if cell.value == "Всього:":
                    isDone = True
                    break

                if cell.value == regionName:
                    returnValueFromTkeDK = tkeDKsheet.cell(column=columnDebtInTkeDK, row=cell.row).value
                    isDone = True
                    break
            if isDone:
                break
        
        if returnValueFromTke == None:
            returnValueFromTke = 0
        if returnValueFromTkeDK == None:
            returnValueFromTkeDK = 0

        return [returnValueFromTke, returnValueFromTkeDK]

    def teplosetiColumnsFromCToF(self, tkeSheet, regionName, fromWhatColumn: str):
        """
        """
        rangeIterInTke = "C15" + ":" + "C" + str(tkeSheet.max_row)
        columnWithData = openpyxl.utils.column_index_from_string(fromWhatColumn)
        returnValue = 0
        for cells in tkeSheet[rangeIterInTke]:
            for cell in cells:
                # If null then its end of data
                if cell.value == None:
                    return 0
                
                if cell.value == regionName:
                    returnValue = tkeSheet.cell(column=columnWithData, row=cell.row).value
                    if returnValue != None:
                        return returnValue
                    elif returnValue == None:
                        return 0
        return 0

    def promColumnB(self, promPrevSheet, promDkSheet, regionName):
        """
        """
        rangeIterInPromPrev = "C12" + ":" + "C" + str(promPrevSheet.max_row)
        columnTotalDebt = openpyxl.utils.column_index_from_string("IG")
        returnValuePromPrev = 0
        isDone = False
        for cells in promPrevSheet[rangeIterInPromPrev]:
            for cell in cells:
                # If cell equal to "Всього:" then it is the end of the sheet
                if cell.value == "Всього:":
                    isDone = True
                    break
                # Add debt of all regions in sheet
                if cell.value == regionName:
                    returnValuePromPrev += promPrevSheet.cell(column=columnTotalDebt, row=cell.row).value
            
            if isDone:
                break
        
        rangeIterInPromDK = "C12" + ":" + "C" + str(promDkSheet.max_row)
        returnValuePromDK = 0
        isDone = False
        for cells in promDkSheet[rangeIterInPromDK]:
            for cell in cells:
                # If cell equal to "Всього:" then it is the end of the sheet
                if cell.value == "Всього:":
                    isDone = True
                    break
                # Add debt of all regions in sheet
                if cell.value == regionName:
                    returnValuePromDK += promDkSheet.cell(column=columnTotalDebt, row=cell.row).value
            
            if isDone:
                break

        return [returnValuePromPrev, returnValuePromDK]

    def promColumnsFromCToF(self, promSheet, region, fromWhatColumn: str):
        """
        """
        rangeIterInProm = "H12" + ":" + "H" + str(promSheet.max_row)
        columnWithData = openpyxl.utils.column_index_from_string(fromWhatColumn)
        returnValue = 0
        for cells in promSheet[rangeIterInProm]:
            for cell in cells:                
                if cell.value == region:
                    returnValue += promSheet.cell(column=columnWithData, row=cell.row).value
        
        if returnValue != None:
            return returnValue
        elif returnValue == None:
            return 0

    def deleteCompanyData(self, promSheet, listOfSpecificCompanies):
        """
        """
        rangeIter = "E12" + ":" + "E" + str(promSheet.max_row)
        for cells in promSheet[rangeIter]:
            for cell in cells:
                if str(cell.value) in listOfSpecificCompanies:
                    
                    for i in range(1, promSheet.max_column):
                        promSheet.cell(column=i, row=cell.row).value = None