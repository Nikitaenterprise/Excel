from src.alg import *
from src.saldo import *

class NKREKU2(Algorithm):

    def checkIfDirectoryIsReady(self, path: str):
        self.mng.addFilesInDir()

        self.template = self.mng.getFile("Шаблон", extension=".xlsx")
        self.template.shouldBeDeleted = False

        self.mng.getFile("Оборотно-сальдова вiдомiсть", 
                        exactMatch=True)
        self.mng.getFile("Оборотно-сальдова вiдомiсть пром", 
                        exactMatch=True)

        self.mng.deleteUnCalledFiles()               
        self.mng.allFromXlsToXlsx()

        try:
            self.saldo = self.mng.getFile(
                                    "Оборотно-сальдова вiдомiсть",
                                    extension=".xlsx", 
                                    exactMatch=True)
            self.saldoEE = self.mng.getFile(
                                    "Оборотно-сальдова вiдомiсть пром", 
                                    extension=".xlsx", 
                                    exactMatch=True)
            
            if self.mng.getNumberOfFiles() != self.numberOfFilesToStart:
                raise AttributeError
        except AttributeError:
            print("Не хватает файлов для работы. Проверьте директорию " + str(path))
            msg = r"""Файлы, нужные для работы:
            1. Шаблон... файл-форма отчетности НКРЭКУ №2 с заполненными 
                        данными в колонке L в строках 48(4.1)-51(4.4)
                    даные заполняются из : 
                    Финансы\Движение денежных средств\Импорт платежей
            2. Оборотно-сальдова вiдомiсть : 
                    за предыдущий месяц (1 - 30 числа),
                    по категориям ТЕ, БО, КП, РО, НС, ВТЕ,
                    без лимитов, без судовых решений, 
                    без ВАТ и ЗБУТ (признак суб.(-) Усі ВАТ+ЗБУТ)
                    (Менеджер отчетов\Стан розрахунків\
                        Оборотно-сальдовая\
                            Оборотно-сальдова відомість... (2gv))
            2. Оборотно-сальдова вiдомiсть пром : 
                    за предыдущий месяц (1 - 30 числа),
                    по категориям промисловість,
                    без лимитов, без судовых решений, 
                    без ВАТ и ЗБУТ (признак суб.(-) Усі ВАТ+ЗБУТ)
                    (Менеджер отчетов\Стан розрахунків\
                        Оборотно-сальдовая\
                            Оборотно-сальдова відомість... (2gv))
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
                self.saldoEE.close()
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
        self.saldoEE.open(data_only=True)
        self.saldoEEWs = self.saldoEE.getWs()
        
        isReady = self.checkIfFileIsReady()
        if isReady:
            self.columnFiller()
            # Add EE to calculation
            # Will be added into "ТКЕ ЕЕ" row
            self.addEE()
            # Save file
            self.saveData()

        else:
            msg = r"""Заполните столбец L с названием :
                перераховано коштів на поточний рахунок із 
                спеціальним режимом використання
            в файле Шаблон.xlsx
            Заполняется из : 
                    Финансы\Движение денежных средств\Импорт платежей
            Фильтры:
                    ТКЕ населення -> Ощадбанк ТКЕ-НАС
                    ТКЕ релігія -> Ощадбанк ТКЕ-РО
                    ТКЕ бюджет -> Укргазбанк ТКЕ-БО
                    ТКЕ ВТЕ, НС, КП -> Укргазбанк ПРЛ
                    ТКЕ ЕЕ -> Укргазбанк Генер ПКМУ-670-Р

            После заполнения данных и сохранения файла 
            запустите программу еще раз"""
            print(msg)
        
        return        

    def checkIfFileIsReady(self):
        
        tempTemplate = self.template
        self.template = self.mng.addFileByPath(
                                            self.template.pathToFile, 
                                            self.template.fileName,
                                            returnFile=True) 
        self.template.open(data_only=True)
        templateWs = self.template.getWs()

        # Get values from pre-writed data in template xlsx
        # from column L
        columnL = openpyxl.utils.column_index_from_string("L")
        startRow = 48
        self.listOfValuesFromColumnL = []
        for i in range(0, 5):
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
                                            ["G", "H", "I", "L"])
            # Multiply by 1000 because "H" column 
            # already divided by 1000 in saldo excel data
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

    def addEE(self):
        # Open saldo with prom and EE contracts 
        # and delete all data except EE contacts
        rangeIter = "A10" + ":" + "A" + str(self.saldoEEWs.max_row)
        for cells in self.saldoEEWs[rangeIter]:
            for cell in cells:
                if cell.value != None and "ЕЕ" not in cell.value:
                    for i in range(1, self.saldoEEWs.max_column):
                        self.saldoEEWs.cell(row=cell.row, 
                                            column=i).value = None
        
        columnToWriteList = ["E", "F", "G", "M", "N", "O", "Q"]
        columnList = []
        for column in columnToWriteList:
            columnList.append(openpyxl.utils.column_index_from_string(column))

        # Get list of values (columns G, H, I, L) from saldo
        toWrite = findInSaldoAllValues(self.saldoEEWs,
                                        None,
                                        None,
                                        ["G", "H", "I", "L"])
        
        # Divide all values by 1000 except data from H 
        for i in range(0, 4):
                toWrite[i] /= 1000
        toWrite[1] *= 1000

        # Get data from T column from saldo. Result is a list
        # with one element
        toN = findInSaldoAllValues(self.saldoEEWs, 
                                        None, 
                                        None, 
                                        ["T"])
        # Subtract  
        # data from column L and row 52 from template sheet
        #             +
        # data from saldo sheet column L
        toN[0] -= (self.listOfValuesFromColumnL[4] + toWrite[3])*1000
        toN[0] /= 1000
        
        # Get data from columns T and U from saldo, then divide by 1000
        toOandQ = findInSaldoAllValues(self.saldoEEWs, 
                                        None, 
                                        ["!2019"], 
                                        ["T", "U"])
        toOandQ[0] /= 1000
        toOandQ[1] /= 1000
        
        toWrite.extend(toN)
        toWrite.extend(toOandQ)
        # Write values into template sheet
        for i,j in zip(columnList, toWrite):
            self.templateWs.cell(column=i, 
                                    row=52).value = j
        
    def saveData(self):
        """Save file with formulas and then reopen it for make 
        it with numbers
        """
        saveName = "НКРЕКП №2 с формулами"
        # Save file with formulas
        self.template.save(self.template.pathToFile, 
                                saveName, 
                                extension=".xlsx")
        # Open this file making it with values only
        newFile = self.mng.addFileByPath(self.template.pathToFile,
                                            saveName+".xlsx",
                                            returnFile=True,
                                            defaultParser=False,
                                            openBy=1)
        newFile.open()

        # Copy and paste data as values
        ws = newFile.getWs(isActiveSheet=True)
        row = 53
        ws.Rows(row).Copy()
        xlPasteValues = -4163
        ws.Rows(row).PasteSpecial(Paste=xlPasteValues)

        # Delete 2 rows with "ТКЕ ВТЕ, НС, КП" and "ТКЕ ЕЕ"
        deleteRows = [51, 51]   # rows 51 and 52 
                                # but after first delete 
                                # second row moves up
        for row in deleteRows:
            ws.Rows(row).Delete()

        # Save file again
        saveName = "НКРЕКП №2"
        newFile.save(self.template.pathToFile, 
                                saveName, 
                                extension=".xlsx")