from src.alg import *

class NKREKU2(Algorithm):

    def checkIfDirectoryIsReady(self, path: str):
        self.mng.addFilesInDir()

        self.outFile = self.mng.getFile("Ф2", extension=".xlsx")
        self.outFile.shouldBeDeleted = False

        self.mng.getFile("Оборотно-сальдова вiдомiсть")

        self.mng.deleteUnCalledFiles()               
        self.mng.allFromXlsToXlsx()

        try:
            self.saldo = self.mng.getFile("Оборотно-сальдова вiдомiсть", extension=".xlsx")
            
            if self.mng.getNumberOfFiles() != self.numberOfFilesToStart:
                raise AttributeError
        except AttributeError:
            print("Не хватает файлов для работы. Проверьте директорию " + str(path))
            msg = r"""Файлы, нужные для работы:
            1. Ф2... файл-форма отчетности НКРЭКУ №2 с заполненными данными в колонке L в строках 48(4.1)-51(4.4)
                    даные заполняются из : Финансы\Движение денежных средств\Импорт платежей
            2. Оборотно-сальдова вiдомiсть : за предыдущий месяц (1 число месяца - последнее число месяца),
                    по категориям ТЕ, БО, КП, РО, НС, ВТЕ,
                    без лимитов, без судовых решений, 
                    без ВАТ и ЗБУТ (признак суб.(-) Усі ВАТ+ЗБУТ)
                    (Менеджер отчетов\Стан розрахунків\Оборотно-сальдовая\Оборотно-сальдова відомість... (2gv))
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
        self.getValues()
        self.deleteFiles()

    def getValues(self):

        self.outFile.open(data_only=False)
        self.outFileWs = self.outFile.getWs()
        self.saldo.open(data_only=True)
        self.saldoWs = self.saldo.getWs()
        
        isReady = self.checkIfFileIsReady()
        if isReady:
            self.columnFiller()
            self.outFile.save(self.outFile.pathToFile, "НКРЕКП №2", extension=".xlsx")
        else:
            msg = r"""Заполните столбец L с названием :
            перераховано коштів на поточний рахунок із спеціальним режимом використання
            Заполняется из : Финансы\Движение денежных средств\Импорт платежей
            Фильтры:
                    4.1 -> Ощадбанк ТКЕ-НАС
                    4.2 -> Ощадбанк ТКЕ-РО
                    4.3 -> Укргазбанк ТКЕ-БО
                    4.4 -> Укргазбанк ПРЛ
            После заполнения данных и сохранения файла запустите программу еще раз"""
            print(msg)
        
        return        

    def checkIfFileIsReady(self):
        
        tempOutFile = self.outFile
        self.outFile = self.mng.addFileByPath(
                                            self.outFile.pathToFile, 
                                            self.outFile.fileName,
                                            returnFile=True
                                            ) 
        self.outFile.open(data_only=True)
        outFileWs = self.outFile.getWs()

        columnL = openpyxl.utils.column_index_from_string("L")
        startRow = 48
        self.listOfValuesFromColumnL = []
        for i in range(0, 4):
            value = outFileWs.cell(column=columnL, row=startRow+i).value
            self.listOfValuesFromColumnL.append(value)

        self.outFile.close()
        self.mng.removeFile(self.outFile)
        self.outFile = tempOutFile

        # If cells are empty then file isn`t ready
        if None in self.listOfValuesFromColumnL:
            return False
        else:
            return True

    def columnFiller(self):
        startRow = 48
        listOfCategories = [
                            ["ТЕ теплоенергетика"], 
                            ["РО теплоенергетика"], 
                            ["БО теплоенергетика"], 
                            ["НС теплоенергетика","КП теплоенергетика", "ВТЕ теплоенергетика"]
                            ]
        
        columnE = openpyxl.utils.column_index_from_string("E")
        for i in range(0, 4):
            toWrite = self.findInSaldoAllValues(self.saldoWs, listOfCategories[i], None, "G")
            self.outFileWs.cell(column=columnE, row=startRow+i).value = toWrite / 1000
        
        columnM = openpyxl.utils.column_index_from_string("M")
        for i in range(0, 4):
            toWrite = self.findInSaldoAllValues(self.saldoWs, listOfCategories[i], None, "L")
            self.outFileWs.cell(column=columnM, row=startRow+i).value = toWrite / 1000
        
        columnO = openpyxl.utils.column_index_from_string("O")
        for i in range(0, 4):
            allResources = self.findInSaldoAllValues(self.saldoWs, listOfCategories[i], None, "T")
            resources2019 = self.findInSaldoAllValues(self.saldoWs, listOfCategories[i], ["2019"], "T")
            toWrite = allResources - resources2019
            self.outFileWs.cell(column=columnO, row=startRow+i).value = toWrite / 1000
        
        columnQ = openpyxl.utils.column_index_from_string("Q")
        for i in range(0, 4):
            allResources = self.findInSaldoAllValues(self.saldoWs, listOfCategories[i], None, "U")
            resources2019 = self.findInSaldoAllValues(self.saldoWs, listOfCategories[i], ["2019"], "U")
            toWrite = allResources - resources2019
            self.outFileWs.cell(column=columnQ, row=startRow+i).value = toWrite / 1000

        columnN = openpyxl.utils.column_index_from_string("N")
        for i in range(0, 4):
            valueFromL = self.listOfValuesFromColumnL[i] * 1000
            valueFromM = self.outFileWs.cell(column=columnM, row=startRow+i).value * 1000
            valueFromSaldo = self.findInSaldoAllValues(self.saldoWs, listOfCategories[i], None, "T")
            toWrite = valueFromSaldo - valueFromL - valueFromM
            self.outFileWs.cell(column=columnN, row=startRow+i).value = toWrite / 1000
        
        return

        
    def findInSaldoAllValues(self, saldoSheet, whatCategory: list, whatResource: list, whatColumn: str):
        """
        """
        rangeIter = "A10" + ":" + "A" + str(saldoSheet.max_row)
        columnCategory = openpyxl.utils.column_index_from_string("C")
        columnResource = openpyxl.utils.column_index_from_string("F")
        columnWithData = openpyxl.utils.column_index_from_string(whatColumn)
        returnValue = 0
        for cells in saldoSheet[rangeIter]:
            for cell in cells:
                
                category = saldoSheet.cell(column=columnCategory,
                                        row=cell.row).value
                # Transform int value of resource into str 2019 -> "2019"
                resource = str(saldoSheet.cell(column=columnResource,
                                        row=cell.row).value).strip()
                
                # Next row if this is name of company
                if category == None and resource == "":
                    continue

                value = saldoSheet.cell(column=columnWithData,
                                        row=cell.row).value
                if value != None:
                    # If both category and resource 
                    # are not specified then add all values
                    if not whatCategory and not whatResource:       
                        returnValue += value
                        continue
                    # If category is specified and resource aren`t
                    if whatCategory and not whatResource:
                        if category in whatCategory:
                            returnValue += value
                            continue
                    # If resource is specified and category aren`t
                    if whatResource and not whatCategory:
                        if resource in whatResource:
                            returnValue += value
                            continue
                    # If both are specified
                    if whatCategory and whatResource:
                        if category in whatCategory and resource in whatResource:
                            returnValue += value
                            continue
        try:
            returnValue
        except UnboundLocalError:
            returnValue = 0
            
        return returnValue