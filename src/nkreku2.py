from src.alg import *
from src.saldo import *

class NKREKU2(Algorithm):

    def checkIfDirectoryIsReady(self, path: str):
        self.mng.addFilesInDir()

        self.template = self.mng.getFile("Шаблон", extension=".xlsx")
        self.template.shouldBeDeleted = False

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
            1. Шаблон... файл-форма отчетности НКРЭКУ №2 с заполненными данными в колонке L в строках 48(4.1)-51(4.4)
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

        self.template.open(data_only=False)
        self.templateWs = self.template.getWs()
        self.saldo.open(data_only=True)
        self.saldoWs = self.saldo.getWs()
        
        isReady = self.checkIfFileIsReady()
        if isReady:
            self.columnFiller()
            self.template.save(self.template.pathToFile, "НКРЕКП №2", extension=".xlsx")
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
        
        tempTemplate = self.template
        self.template = self.mng.addFileByPath(
                                            self.template.pathToFile, 
                                            self.template.fileName,
                                            returnFile=True
                                            ) 
        self.template.open(data_only=True)
        templateWs = self.template.getWs()

        columnL = openpyxl.utils.column_index_from_string("L")
        startRow = 48
        self.listOfValuesFromColumnL = []
        for i in range(0, 4):
            value = templateWs.cell(column=columnL, row=startRow+i).value
            self.listOfValuesFromColumnL.append(value)

        self.template.close()
        self.mng.removeFile(self.template)
        self.template = tempTemplate

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
                                [
                                "НС теплоенергетика",
                                "КП теплоенергетика", 
                                "ВТЕ теплоенергетика"
                                ]
                            ]
        columnToWriteList = ["E", "F", "G", "M"]
        columnList = []
        for column in columnToWriteList:
            columnList.append(openpyxl.utils.column_index_from_string(column))

        for i in range(0, 4):
            toWrite = findInSaldoAllValues(self.saldoWs, 
                                            listOfCategories[i],
                                            None,
                                            ["G", "H", "I", "L",])
            # Multiply by 1000 because its already divided by 1000 
            # in saldo excel data
            toWrite[1] *= 1000

            for j in range(0, 4):
                self.templateWs.cell(column=columnList[j], 
                            row=startRow+i).value = toWrite[j] / 1000
            
        columnO = openpyxl.utils.column_index_from_string("O")
        columnQ = openpyxl.utils.column_index_from_string("Q")
        for i in range(0, 4):
            toWrite = findInSaldoAllValues(self.saldoWs, 
                                            listOfCategories[i], 
                                            ["!2019"], 
                                            ["T", "U"])
            self.templateWs.cell(column=columnO, row=startRow+i).value =\
                        toWrite[0] / 1000
            self.templateWs.cell(column=columnQ, row=startRow+i).value =\
                        toWrite[1] / 1000
        
        columnN = openpyxl.utils.column_index_from_string("N")
        columnM = openpyxl.utils.column_index_from_string("M")
        for i in range(0, 4):
            valueFromL = self.listOfValuesFromColumnL[i] * 1000
            valueFromM = self.templateWs.cell(column=columnM, 
                                        row=startRow+i).value * 1000
            valueFromSaldo = findInSaldoAllValues(self.saldoWs, 
                                        listOfCategories[i], 
                                        None, 
                                        ["T"])
            toWrite = valueFromSaldo[0] - valueFromL - valueFromM
            self.templateWs.cell(column=columnN, row=startRow+i).value =\
                                        toWrite / 1000

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
        try:
            returnValue
        except UnboundLocalError:
            returnValue = 0
            
        return returnValue