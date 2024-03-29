from src.alg import *
from src.saldo import *

class NKREKU3(Algorithm):

    def checkIfDirectoryIsReady(self, path: str):
        self.mng.addFilesInDir()

        self.template = self.mng.getFile("Шаблон", extension=".xlsx")
        self.template.shouldBeDeleted = False

        self.mng.getFile("Оборотно-сальдова вiдомiсть", exactMatch=True)
        self.mng.getFile("Оборотно-сальдова вiдомiсть пром", exactMatch=True)

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
            1. Шаблон... файл-форма отчетности НКРЭКУ №3
            2. Оборотно-сальдова вiдомiсть : 
                    за предыдущий месяц (1 - 30),
                    по категориям 
                            ТЕ, РО, КП, БО, НС, ВТЕ
                    без лимитов, 
                    без судовых решений, 
                    без ВАТ и ЗБУТ (признак суб.(-) Усі ВАТ+ЗБУТ)
                    (Менеджер отчетов\Стан розрахунків\
                                Оборотно-сальдовая\
                                Оборотно-сальдова відомість... (2gv))
            2. Оборотно-сальдова вiдомiсть пром : 
                    за предыдущий месяц (1 - 30),
                    по категориям промисловість,
                    без лимитов, 
                    без судовых решений, 
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
        self.fill()
        self.template.save(self.template.pathToFile, 
                            "НКРЕКП №3",
                            extension=".xlsx")
        self.deleteFiles()

    def fill(self):

        self.template.open(data_only=False)
        self.templateWs = self.template.getWs()
        self.saldo.open(data_only=True)
        self.saldoWs = self.saldo.getWs()
        self.saldoEE.open(data_only=True)
        self.saldoEEWs = self.saldoEE.getWs()

        startRow = 45
        columnD = openpyxl.utils.column_index_from_string("D")
        columnE = openpyxl.utils.column_index_from_string("E")
        listOfCat = [
                        ["ТЕ теплоенергетика"],
                        ["РО теплоенергетика"]
                    ]
        costList = []
        for category in listOfCat:
            tkeCategory = findInSaldoAllValues(self.saldoWs,
                                            category,
                                            None,
                                            ["H", "I"])
            TkePrice = 0
            if tkeCategory[0] != 0 and tkeCategory[0] != None:
                TkePrice = tkeCategory[1] / tkeCategory[0]
                TkePrice = TkePrice / 1.2 - 124.16
            costList.append(TkePrice)
            self.templateWs.cell(column=columnD, 
                                    row=startRow).value = tkeCategory[0]
            self.templateWs.cell(column=columnE, 
                                    row=startRow).value = TkePrice
            startRow += 1
        
        # Open saldo with prom and EE contracts 
        # and delete all data except EE contacts
        rangeIter = "A10" + ":" + "A" + str(self.saldoEEWs.max_row)
        for cells in self.saldoEEWs[rangeIter]:
            for cell in cells:
                if cell.value != None and "ЕЕ" not in cell.value:
                    for i in range(1, self.saldoEEWs.max_column):
                        self.saldoEEWs.cell(row=cell.row, 
                                            column=i).value = None
        # Calculate other categories 
        listOfCat = [
                        "БО теплоенергетика",
                        "КП теплоенергетика",
                        "НС теплоенергетика",
                        "ВТЕ теплоенергетика"
                    ]
        otherTke = findInSaldoAllValues(self.saldoWs,
                                            listOfCat,
                                            None,
                                            ["H", "I"])
        # Calculate EE
        tkeEe = findInSaldoAllValues(self.saldoEEWs,
                                        None,
                                        None,
                                        ["H", "I"])
        # Add EE data to other
        otherTke[0] += tkeEe[0]
        otherTke[1] += tkeEe[1]

        TkeOthersPrice = 0
        if otherTke[0] != 0 and otherTke[0] != None:
            TkeOthersPrice = otherTke[1] / otherTke[0]
            TkeOthersPrice = TkeOthersPrice / 1.2 - 124.16
        costList.append(TkeOthersPrice)
        self.templateWs.cell(column=columnD, 
                            row=startRow).value = otherTke[0]
        self.templateWs.cell(column=columnE, 
                            row=startRow).value = TkeOthersPrice

        for i in range(0, len(costList)-1):
            delta = costList[i+1] - costList[i]
            if delta >= 0.0001 or delta <= -0.0001:
                print("Цены не сходятся")