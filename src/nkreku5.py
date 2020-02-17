from src.alg import *
from src.saldo import *

class NKREKU5(Algorithm):

    def checkIfDirectoryIsReady(self, path:str):
        self.mng.addFilesInDir()

        self.template = self.mng.getFile("Шаблон", extension=".xlsx")
        self.template.shouldBeDeleted = False

        self.mng.getFile("Оборотно-сальдова вiдомiсть 1", 
                        exactMatch=True)
        self.mng.getFile("Оборотно-сальдова вiдомiсть 2", 
                        exactMatch=True)
        self.mng.getFile("Оборотно-сальдова вiдомiсть 3", 
                        exactMatch=True)
        self.mng.getFile("Оборотно-сальдова вiдомiсть 3 месяца", 
                        exactMatch=True)
        
        self.mng.deleteUnCalledFiles()               
        self.mng.allFromXlsToXlsx()

        try:
            self.saldoFirstMonth = self.mng.getFile(
                                    "Оборотно-сальдова вiдомiсть 1",
                                    extension=".xlsx", 
                                    exactMatch=True)
            self.saldoSecondMonth = self.mng.getFile(
                                    "Оборотно-сальдова вiдомiсть 2",
                                    extension=".xlsx", 
                                    exactMatch=True)
            self.saldoThirdMonth = self.mng.getFile(
                                    "Оборотно-сальдова вiдомiсть 3",
                                    extension=".xlsx", 
                                    exactMatch=True)
            self.saldoTotal = self.mng.getFile(
                                    "Оборотно-сальдова вiдомiсть 3 месяца",
                                    extension=".xlsx", 
                                    exactMatch=True)
            if self.mng.getNumberOfFiles() != self.numberOfFilesToStart:
                raise AttributeError
        except AttributeError:
            print("Не хватает файлов для работы. Проверьте директорию " + str(path))
            msg = r"""Файлы, нужные для работы:
            1. Шаблон : файл-форма отчетности НКРЭКУ №5
            2. Оборотно-сальдова вiдомiсть 3 месяца : 
                        за квартал :
                            1 число первого месяца периода - 
                            30 число последнего месяца периода
                            например, 1 июля - 30 сентября
                        по категориям :
                            ТЕ, 
                            БО, 
                            КП, 
                            РО, 
                            НС, 
                            ВТЕ,
                            промисловість
                        признак субъекта :
                            (-) Усі ВАТ+ЗБУТ
                        без лимитов, 
                        без судовых решений
            
            3. Оборотно-сальдова вiдомiсть 1
                        за первый месяц квартала :
                            1 число - 30 число месяца
                        --//--
            4. Оборотно-сальдова вiдомiсть 2
                        за второй месяц квартала :
                            1 число - 30 число месяца
                        --//--
            5. Оборотно-сальдова вiдомiсть 3
                        за третий месяц квартала :
                            1 число - 30 число месяца
                        --//--

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
                self.saldoTotal.close()
                self.saldoFirstMonth.close()
                self.saldoSecondMonth.close()
                self.saldoThirdMonth.close()
            except:
                print(bcolors.WARNING +\
                    "Программа не смогла закрыть экселевские файлы"\
                    + bcolors.ENDC)
        self.mng.deleteClosedFiles()
        return

    def run(self):
        self.template.open(data_only=False)
        self.templateWs = self.template.getWs()
        
        self.saldoFirstMonth.open(data_only=True)
        self.saldoFirstMonthWs = self.saldoFirstMonth.getWs()
        self.saldoSecondMonth.open(data_only=True)
        self.saldoSecondMonthWs = self.saldoSecondMonth.getWs()
        self.saldoThirdMonth.open(data_only=True)
        self.saldoThirdMonthWs = self.saldoThirdMonth.getWs()
        self.saldoTotal.open(data_only=True)
        self.saldoTotalWs = self.saldoTotal.getWs()

        self.listWithSaldos = [self.saldoFirstMonth,
                                self.saldoSecondMonth,
                                self.saldoThirdMonth
                                ]
        self.listWithSaldoWs = [self.saldoFirstMonthWs,
                                self.saldoSecondMonthWs,
                                self.saldoThirdMonthWs
                                ]

        self.section1()
        self.section2()

        self.template.save(self.template.pathToFile,
                            "НКРЕКП №5",
                            extension=".xlsx")
        self.deleteFiles()
    
    def section1(self):
        listOfCategories = [
                            [
                                "ТЕ теплоенергетика",
                                "ТЕ (газовий депозит)"
                            ], 
                            ["БО теплоенергетика"], 
                            ["РО теплоенергетика"], 
                                [
                                "НС теплоенергетика",
                                "КП теплоенергетика", 
                                "ВТЕ теплоенергетика"
                                ],
                            ["промисловість"]
                            ]
        columnList = [columnIndexFromString(x) for x in ["G", "H", "I"]]
        startRow = 43
        # Iterate in 3 months (3 groups of columns)
        for i in range (0, len(self.listWithSaldoWs)):
            # Iterate in categories
            for j in range(0, len(listOfCategories)):
                dataFromSaldo = findInSaldoAllValues(self.listWithSaldoWs[i],
                                                    listOfCategories[j],
                                                    None,
                                                    ["H", "I"])
                costWithPDV = dataFromSaldo[1] / dataFromSaldo[0]
                costWithoutPDV = costWithPDV / 1.2
                # Group in one list
                dataList = [dataFromSaldo[0], 
                            costWithoutPDV, 
                            costWithPDV
                            ]
                # Iterate in one column group in 3 columns
                for k in range(0, 3):
                    self.templateWs.cell(column=columnList[k]+3*i,
                                    row=startRow+j).value = dataList[k]


    def section2(self):
        listOfCategories = [
                            [
                                "ТЕ теплоенергетика",
                                "ТЕ (газовий депозит)"
                            ], 
                            ["БО теплоенергетика"], 
                            ["РО теплоенергетика"], 
                            [
                                "НС теплоенергетика",
                                "КП теплоенергетика", 
                                "ВТЕ теплоенергетика"
                            ],
                            ["промисловість"]
                            ]
        columnList = [columnIndexFromString(x) for x in ["E", "F", 
                                                         "G", "J", "H"]]
        startRow = 61
        for i in range(0, len(listOfCategories)):

            dataFromSaldo = findInSaldoAllValues(self.saldoTotalWs,
                                                listOfCategories[i],
                                                None,
                                                ["H", "I", "T", "U"])
            forPreviousYears = findInSaldoAllValues(self.saldoTotalWs,
                                                listOfCategories[i],
                                                ["!2019", "!2020"],
                                                ["T"])

            # Group data in one list for simplisity
            # Data grouped in right order according 
            # to columnList order
            # also data is divided by 1000 except first one
            dataFromSaldo[0] *= 1000
            dataList = []
            for element in dataFromSaldo:
                dataList.append(element/1000)
            for element in forPreviousYears:
                dataList.append(element/1000)

            # Iterate in columns in one row
            for j in range(0, len(columnList)):
                self.templateWs.cell(column=columnList[j], 
                                    row=startRow+i).value = dataList[j]


