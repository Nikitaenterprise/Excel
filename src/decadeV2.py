from src.alg import *
from src.saldo import *
import time

class Decade(Algorithm):

    def checkIfDirectoryIsReady(self, path:str):
        self.mng.addFilesInDir()

        self.template = self.mng.getFile("Шаблон", 
                                    extension=".xlsx")
        self.template.shouldBeDeleted = False
        self.tkeDK = self.mng.getFile("ТКЕ ДК", 
                                    extension=".xlsx", exactMatch=True)
        self.tkeDK.shouldBeDeleted = False
        self.promDK = self.mng.getFile("Промисловість_ ДК", 
                                    extension=".xlsx", exactMatch=True)
        self.promDK.shouldBeDeleted = False


        self.mng.getFile("Оборотно-сальдова вiдомiсть", 
                                    exactMatch=True)
        # self.mng.getFile("Оборотно-сальдова вiдомiсть последний месяц", 
        #                             exactMatch=True)

        self.mng.deleteUnCalledFiles()               
        self.mng.allFromXlsToXlsx()

        try:
            self.saldo = self.mng.getFile("Оборотно-сальдова вiдомiсть", 
                                    extension=".xlsx", exactMatch=True)
            # self.saldoLastMonth = self.mng.getFile("Оборотно-сальдова вiдомiсть последний месяц", 
            #                         extension=".xlsx", exactMatch=True)
            
            if self.mng.getNumberOfFiles() != self.numberOfFilesToStart:
                    raise AttributeError
        except AttributeError:
            print("Не хватает файлов для работы. Проверьте директорию " + str(path))
            msg = r"""Файлы, нужные для работы : """
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
                self.saldoLastMonth.close()
            except:
                print(bcolors.WARNING +\
                    "Программа не смогла закрыть экселевские файлы"\
                    + bcolors.ENDC)
        self.mng.deleteClosedFiles()

    def run(self):
        # self.naselenie()
        # self.religion()
        # self.budget()
        self.teploseti()
        # self.forPresident()
        # self.generations()
        self.template.save(self.template.pathToFile, 
                            "На печать", extension=".xlsx")
        self.deleteFiles()

    def naselenie(self):
        self.template.open(data_only=False)
        templateWsNas = self.template.getWs("Населення")
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()
        #self.saldoLastMonth.open(data_only=True)
        #saldoLastMonthWs = self.saldoLastMonth.getWs()

        self.fill(templateWs=templateWsNas,
                        saldoWs=saldoWs,
                        listOfCategories=[
                                        "населення", 
                                        "населення (газовий депозит)"
                                        ])

    def religion(self):
        self.template.open(data_only=False)
        templateWsReligion = self.template.getWs("Релігія")
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()
        # self.saldoLastMonth.open(data_only=True)
        # saldoLastMonthWs = self.saldoLastMonth.getWs()

        self.fill(templateWs=templateWsReligion,
                        saldoWs=saldoWs,
                        listOfCategories=[
                                        "релігійні організації", 
                                        "вічний вогонь"
                                        ])

    def budget(self):
        self.template.open(data_only=False)
        templateWsBudget = self.template.getWs("Бюджет")
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()
        # self.saldoLastMonth.open(data_only=True)
        # saldoLastMonthWs = self.saldoLastMonth.getWs()

        self.fill(templateWs=templateWsBudget,
                        saldoWs=saldoWs,
                        listOfCategories=[
                                        "бюджет"
                                        ])

    def teploseti(self):
        start = time.time()
        self.template.open(data_only=False)
        templateWsTeploseti = self.template.getWs("Тепломережі")
        print("opened template =", time.time()-start)
        start = time.time()
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()
        print("opened TKE =", time.time()-start)
        start = time.time()
        self.tkeDK.open(data_only=True)
        tkeDKWs = self.tkeDK.getWs()
        print("opened DK =", time.time()-start)

        self.fill(templateWs=templateWsTeploseti,
                        saldoWs=saldoWs,
                        listOfCategories=[
                                        "ТЕ теплоенергетика",
                                        "ТЕ (газовий депозит)",
                                        "БО теплоенергетика",
                                        "РО теплоенергетика",
                                        "НС теплоенергетика",
                                        "КП теплоенергетика",
                                        "ВТЕ теплоенергетика"
                                        ],
                        dkSheet=tkeDKWs)
        
    def fill(self, templateWs, saldoWs, listOfCategories, dkSheet=None):

        rangeIter = "A9" + ":" + "A" + str(templateWs.max_row-2)
        columnsToFill = [columnIndexFromString(x) for x in 
                            ["B", "C", "D", "E", "F", "H", "I"]]
        
        # If calculating TKE then set DkTkeDept to true
        DkTkeDept, DkPromDept = False, False
        if "ТЕ теплоенергетика" in listOfCategories: 
            DkTkeDept = True
        # If calculating prom then set DkPromDept to true
        elif "промисловість" in listOfCategories:
            DkPromDept = True

        for cells in templateWs[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    allYearsDebt = findInSaldoAllValues(
                                        saldoSheet=saldoWs, 
                                        whatCategory=listOfCategories,
                                        whatResource=None,
                                        whatColumns=["G"],
                                        whatRegion=[cell.value])[0]
                    if DkTkeDept:
                        debtFromDk = self.debtFromTkeDk(region=cell.value,
                                                        dkSheet=dkSheet)
                        allYearsDebt += debtFromDk
                    
                    if DkPromDept:
                        debtFromDk = self.debtFromPromDk(region=cell.value,
                                                        dkSheet=dkSheet)
                        allYearsDebt += debtFromDk

                    previousYearDebt = findInSaldoAllValues(
                                        saldoSheet=saldoWs, 
                                        whatCategory=listOfCategories,
                                        whatResource=["2018"],
                                        whatColumns=["U"],
                                        whatRegion=[cell.value])[0]

                    data = findInSaldoAllValues(
                                        saldoSheet=saldoWs, 
                                        whatCategory=listOfCategories,
                                        whatResource=["2019"],
                                        whatColumns=["H", "I", "K"],
                                        whatRegion=[cell.value])
                    consumedGas = data[0]
                    cost = data[1]
                    payment = data[2]
                    
                    data1 = findInSaldoAllValues(
                                        saldoSheet=saldoWs, 
                                        whatCategory=listOfCategories,
                                        whatResource=None,
                                        whatColumns=["T", "J"],
                                        whatRegion=[cell.value])
                    income = data1[0]
                    penalty = data1[1]

                    data2 = findInSaldoAllValues(
                                        saldoSheet=saldoWs, 
                                        whatCategory=listOfCategories,
                                        whatResource=["2019"],
                                        whatColumns=["T", "U"],
                                        whatRegion=[cell.value])
                    income2020 = data2[0]
                    debt2020 = data2[1]

                    totalDebt = 0
                    totalDebt += allYearsDebt - income
                    totalDebt += income2020 + penalty
                    totalDebt += consumedGas - payment

                    listToPutInTemplate = []
                    listToPutInTemplate.append(allYearsDebt/1000)
                    listToPutInTemplate.append(previousYearDebt/1000)
                    listToPutInTemplate.append(consumedGas/1000)
                    listToPutInTemplate.append(cost/1000)
                    listToPutInTemplate.append(payment/1000)
                    
                    listToPutInTemplate.append(totalDebt/1000)
                    listToPutInTemplate.append(debt2020/1000)

                    # Fill one row in template file
                    for i in range(0, len(listToPutInTemplate)):
                        templateWs.cell(column=columnsToFill[i],
                            row=cell.row).value = listToPutInTemplate[i]

                    
    def debtFromTkeDk(self, region, dkSheet):
        rangeIter = "B12" + ":" + "B" + str(dkSheet.max_row)
        columnDebt = columnIndexFromString("EG")
        debt = 0
        isDone = False
        for cells in dkSheet[rangeIter]:
            for cell in cells:
                if cell.value == "Всього:":
                    isDone = True
                    break

                if cell.value == region:
                    debt = dkSheet.cell(column=columnDebt, 
                                        row=cell.row).value
                    isDone = True
                    break
            if isDone:
                break
        
        if debt == None:
            debt = 0

        return debt
    
    def debtFromPromDk(self, region, dkSheet):
        rangeIter = "C12" + ":" + "C" + str(dkSheet.max_row)
        columnTotalDebt = columnIndexFromString("IG")
        debt = 0
        isDone = False
        for cells in dkSheet[rangeIter]:
            for cell in cells:
                # If cell equal to "Всього:" then it is the end of the sheet
                if cell.value == "Всього:":
                    isDone = True
                    break
                # Add debt of all regions in sheet
                if cell.value == region:
                    value = dkSheet.cell(column=columnTotalDebt, 
                                        row=cell.row).value
                    if value == None:
                        value = 0
                    debt += value
            
            if isDone:
                break
        
        return debt
