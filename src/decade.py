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
        self.mng.getFile("Оборотно-сальдова вiдомiсть последний месяц", 
                                    exactMatch=True)

        self.mng.deleteUnCalledFiles()               
        self.mng.allFromXlsToXlsx()

        try:
            self.saldo = self.mng.getFile("Оборотно-сальдова вiдомiсть", 
                                    extension=".xlsx", exactMatch=True)
            self.saldoLastMonth = self.mng.getFile("Оборотно-сальдова вiдомiсть последний месяц", 
                                     extension=".xlsx", exactMatch=True)
            
            if self.mng.getNumberOfFiles() != self.numberOfFilesToStart:
                    raise AttributeError
        except AttributeError:
            print("Не хватает файлов для работы. Проверьте директорию " + str(path))
            msg = r"""Файлы, нужные для работы : 
            1. Оборотно-сальдова вiдомiсть : 
                                1 января - по декаду,
                                по всем категориям, 
                                с планами
                        (Менеджер отчетов\Стан розрахунків\
                            Оборотно-сальдовая\
                                Оборотно-сальдова відомість... (2gv))
            2. Оборотно-сальдова вiдомiсть последний месяц: 
                                1 января - 30 число предыдущего месяца,
                                по всем категориям 
                        (Менеджер отчетов\Стан розрахунків\
                            Оборотно-сальдовая\
                                Оборотно-сальдова відомість... (2gv))
            Файлы, которые не надо загружать, 
            но они должны быть в папке
            1. Шаблон : пустой шаблон
            2. ТКЕ ДК : база за 2010-2011 год, 
                                построена в 2018 году
                        (Dept\Monitoring.Gas\4. Газ України\
                            ДК Газ України.  База ТКЕ 2010-2011)
            3. Промисловість_ ДК : база за 2010-2011 год, 
                                построена в 2018 году
                        (Dept\Monitoring.Gas\4. Газ України\
                            ДК Газ України. Промисловість)
            После исправления запустите программу заново. 
            Сейчас программа завершит работу
            Нажмите любую клавишу а затем Enter"""
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
        self.getNumberOfDecade()

        self.naselenie()
        self.religion()
        self.budget()
        self.teploseti()
        self.promishlennost()
        self.forPresident()
        self.template.save(self.template.pathToFile, 
                            "На печать", extension=".xlsx")
        self.deleteFiles()

    def getNumberOfDecade(self):
        self.numberOfDecade = 1
        try:
            print("Какая декада?")
            print("1 - первая декада")
            print("2 - вторая декада")
            print("3 - третья декада")
            answer = int(input())
            if(answer == 1 or answer == 2 or answer == 3):
                self.numberOfDecade = answer
        except:
            print("Неправильно введен номер декады")
            print("Программа считает что декада первая")


    def naselenie(self):
        self.template.open(data_only=False)
        templateWs = self.template.getWs("Населення")
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()
        self.saldoLastMonth.open(data_only=True)
        saldoLastMonthWs = self.saldoLastMonth.getWs()
        
        # Should be list because of immutibility of basic types
        self.naselenieForPresident = []
        self.naselenieForPresident.append(0)

        rangeIter = "A9" + ":" + "A" + str(templateWs.max_row-2)
        totalData = []
        for cells in templateWs[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    value = self.getData(region=cell.value,
                            saldoWs=saldoWs,
                            saldoLastMonthWs=saldoLastMonthWs,
                            listOfCategories=[
                                            "населення", 
                                            "населення (газовий депозит)"
                                            ],
                            dataForPresident=self.naselenieForPresident)
                    
                    totalData.append(value)
        
        startRow = 9
        for i in range(0, len(totalData)):
            for j in range(0, len(totalData[i])):
                templateWs.cell(column=columnIndexFromString("B")+j,
                                row=startRow+i).value = totalData[i][j]

    def religion(self):
        self.template.open(data_only=False)
        templateWs = self.template.getWs("Релігія")
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()
        self.saldoLastMonth.open(data_only=True)
        saldoLastMonthWs = self.saldoLastMonth.getWs()

        self.religionForPresident = []
        self.religionForPresident.append(0)

        rangeIter = "A9" + ":" + "A" + str(templateWs.max_row-2)
        totalData = []
        for cells in templateWs[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    value = self.getData(region=cell.value,
                            saldoWs=saldoWs,
                            saldoLastMonthWs=saldoLastMonthWs,
                            listOfCategories=[
                                            "релігійні організації", 
                                            "вічний вогонь"
                                            ],
                            dataForPresident=self.religionForPresident)
                    
                    totalData.append(value)

        startRow = 9
        for i in range(0, len(totalData)):
            for j in range(0, len(totalData[i])):
                templateWs.cell(column=columnIndexFromString("B")+j,
                                row=startRow+i).value = totalData[i][j]

    def budget(self):
        self.template.open(data_only=False)
        templateWs = self.template.getWs("Бюджет")
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()
        self.saldoLastMonth.open(data_only=True)
        saldoLastMonthWs = self.saldoLastMonth.getWs()

        self.budgetForPresident = []
        self.budgetForPresident.append(0)
        
        rangeIter = "A9" + ":" + "A" + str(templateWs.max_row-2)
        totalData = []
        for cells in templateWs[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    value = self.getData(region=cell.value,
                            saldoWs=saldoWs,
                            saldoLastMonthWs=saldoLastMonthWs,
                            listOfCategories=[
                                            "бюджет"
                                            ],
                            dataForPresident=self.budgetForPresident)
                    
                    totalData.append(value)

        startRow = 9
        for i in range(0, len(totalData)):
            for j in range(0, len(totalData[i])):
                templateWs.cell(column=columnIndexFromString("B")+j,
                                row=startRow+i).value = totalData[i][j]

    def teploseti(self):
        self.template.open(data_only=False)
        templateWs = self.template.getWs("Тепломережі")
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()
        self.saldoLastMonth.open(data_only=True)
        saldoLastMonthWs = self.saldoLastMonth.getWs()
        self.tkeDK.open(data_only=True)
        tkeDKWs = self.tkeDK.getWs()

        self.tkeForPresident = []
        self.tkeForPresident.append(0)

        rangeIter = "A9" + ":" + "A" + str(templateWs.max_row-2)
        totalData = []
        for cells in templateWs[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    value = self.getData(region=cell.value,
                            saldoWs=saldoWs,
                            saldoLastMonthWs=saldoLastMonthWs,
                            listOfCategories=[
                                            "ТЕ теплоенергетика",
                                            "ТЕ (газовий депозит)",
                                            "БО теплоенергетика",
                                            "РО теплоенергетика",
                                            "НС теплоенергетика",
                                            "КП теплоенергетика",
                                            "ВТЕ теплоенергетика"
                                            ],
                            dataForPresident=self.tkeForPresident,
                            dkSheet=tkeDKWs)
                    
                    totalData.append(value)

        startRow = 9
        for i in range(0, len(totalData)):
            for j in range(0, len(totalData[i])):
                templateWs.cell(column=columnIndexFromString("B")+j,
                                row=startRow+i).value = totalData[i][j]

    def promishlennost(self):
        self.template.open(data_only=False)
        templateWs = self.template.getWs("Промисловість")
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()
        self.saldoLastMonth.open(data_only=True)
        saldoLastMonthWs = self.saldoLastMonth.getWs()
        self.promDK.open(data_only=True)
        promDKWs = self.promDK.getWs()

        self.promForPresident = []
        self.promForPresident.append(0)

        rangeIter = "A9" + ":" + "A" + str(templateWs.max_row-2)
        totalData = []
        for cells in templateWs[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    value = self.getData(region=cell.value,
                            saldoWs=saldoWs,
                            saldoLastMonthWs=saldoLastMonthWs,
                            listOfCategories=[
                                            "промисловість"
                                            ],
                            dataForPresident=self.promForPresident,
                            dkSheet=promDKWs)
                    
                    totalData.append(value)

        startRow = 9
        for i in range(0, len(totalData)):
            for j in range(0, len(totalData[i])):
                templateWs.cell(column=columnIndexFromString("B")+j,
                                row=startRow+i).value = totalData[i][j]

        # Generations 
        # Create category "генерації" if contract contain EE
        self.generationsCreation(saldoWs)
        self.generationsCreation(saldoLastMonthWs)

        totalData = []
        dontNeedThis = []
        dontNeedThis.append(0)
        for cells in templateWs[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    value = self.getData(region=cell.value,
                            saldoWs=saldoWs,
                            saldoLastMonthWs=saldoLastMonthWs,
                            listOfCategories=[
                                            "генерації"
                                            ],
                            dataForPresident=dontNeedThis)
                    
                    totalData.append(value)

        # Summ all regions in one row
        templateWs = self.template.getWs("Зведена")
        toWrite = []
        for j in range(0, len(totalData[0])):
            value = 0
            for i in range(0, len(totalData)):
                value += totalData[i][j]
            toWrite.append(value)
        
        # Recalculate percent value because it was summed up
        if toWrite[3] != 0:
            toWrite[5] = toWrite[4]/toWrite[3] * 100
        else:
            toWrite[5] = 0
        
        # Fill row with Генерації
        for i in range(0, len(toWrite)):
            templateWs.cell(column=columnIndexFromString("B")+i,
                                row=17).value = toWrite[i]
        
    def getData(self, region, saldoWs, saldoLastMonthWs,
            listOfCategories, dataForPresident, dkSheet=None):

        # If calculating TKE then set DkTkeDept to true
        DkTkeDept, DkPromDept = False, False
        if "ТЕ теплоенергетика" in listOfCategories: 
            DkTkeDept = True
        # If calculating prom then set DkPromDept to true
        elif "промисловість" in listOfCategories:
            DkPromDept = True

        allYearsDebt = findInSaldoAllValues(
                            saldoSheet=saldoWs, 
                            whatCategory=listOfCategories,
                            whatResource=None,
                            whatColumns=["G"],
                            whatRegion=[region])[0]
        if DkTkeDept:
            debtFromDk = self.debtFromTkeDk(region=region,
                                            dkSheet=dkSheet)
            allYearsDebt += debtFromDk
        
        if DkPromDept:
            debtFromDk = self.debtFromPromDk(region=region,
                                            dkSheet=dkSheet)
            allYearsDebt += debtFromDk

        previousYearDebt = findInSaldoAllValues(
                            saldoSheet=saldoWs, 
                            whatCategory=listOfCategories,
                            whatResource=["2019"],
                            whatColumns=["U"],
                            whatRegion=[region])[0]

        data = findInSaldoAllValues(
                            saldoSheet=saldoWs, 
                            whatCategory=listOfCategories,
                            whatResource=["2020"],
                            whatColumns=["H", "I", "T"],
                            whatRegion=[region])
        consumedGas = data[0]
        cost = data[1]
        payment = data[2]

        tmp = findInSaldoAllValues(
                            saldoSheet=saldoLastMonthWs,
                            whatCategory=listOfCategories,
                            whatResource=["2020"],
                            whatColumns=["H", "I"],
                            whatRegion=[region])
        # Gas amount and gas cost fron start  
        # of the year to previous month
        consumedGasLastMonth = tmp[0]
        costLastMonth = tmp[1]
        # Find gas amount and cost only for this month
        consumedGas -= consumedGasLastMonth
        cost -= costLastMonth
        # Divide by number of decade to get 
        # amount of gas in this period
        consumedGas *= self.numberOfDecade/3
        cost *= self.numberOfDecade/3
        # Add this month gas to total
        consumedGas += consumedGasLastMonth
        cost += costLastMonth
        
        data1 = findInSaldoAllValues(
                            saldoSheet=saldoWs, 
                            whatCategory=listOfCategories,
                            whatResource=None,
                            whatColumns=["T", "J"],
                            whatRegion=[region])
        income = data1[0]
        # Write total income to dataForPresident value
        dataForPresident[0] += data1[0]
        penalty = data1[1]

        debt2020 = findInSaldoAllValues(
                            saldoSheet=saldoWs, 
                            whatCategory=listOfCategories,
                            whatResource=["2020"],
                            whatColumns=["U"],
                            whatRegion=[region])[0]

        totalDebt = 0
        #totalDebt += allYearsDebt - income
        #totalDebt += penalty
        #totalDebt += cost - payment

        totalDebt += allYearsDebt + cost + penalty - income

        listOfOneRegion = []
        listOfOneRegion.append(allYearsDebt/1000)
        listOfOneRegion.append(previousYearDebt/1000)
        listOfOneRegion.append(consumedGas/1000)
        listOfOneRegion.append(cost/1000)
        listOfOneRegion.append(payment/1000)
        if cost != 0:
            listOfOneRegion.append(payment/cost*100)
        else:
            listOfOneRegion.append(0)
        listOfOneRegion.append(totalDebt/1000)
        listOfOneRegion.append(debt2020/1000)
        
        return listOfOneRegion

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

    def generationsCreation(self, saldoWs):
        rangeIter = "A10" + ":" + "A" + str(saldoWs.max_row)
        columnCategory = columnIndexFromString("C")
        for cells in saldoWs[rangeIter]:
            for cell in cells:
                if cell.value != None and "ЕЕ" in cell.value:
                    saldoWs.cell(column=columnCategory,
                                row=cell.row).value = "генерації"

    def forPresident(self):
        self.template.open(data_only=False)
        templateWsPivot = self.template.getWs("Зведена")
        columnWhereToWrite = columnIndexFromString("J")
        try:
            templateWsPivot.cell(column=columnWhereToWrite, 
                    row=13).value = self.naselenieForPresident[0] / 1000
            templateWsPivot.cell(column=columnWhereToWrite, 
                    row=14).value = self.budgetForPresident[0] / 1000
            templateWsPivot.cell(column=columnWhereToWrite, 
                    row=15).value = self.tkeForPresident[0] / 1000
            templateWsPivot.cell(column=columnWhereToWrite, 
                    row=16).value = self.promForPresident[0] / 1000
            templateWsPivot.cell(column=columnWhereToWrite, 
                    row=19).value = self.religionForPresident[0] / 1000
        except (UnboundLocalError, AttributeError):
            print("Не заполнен последний столбец во вкладке со сводной таблицей")

                    