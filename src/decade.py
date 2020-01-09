from src.alg import *

class Decade(Algorithm):

    def checkIfDirectoryIsReady(self, path: str):
        self.mng.addFilesInDir()

        # Checks if the files are present
        self.template = self.mng.getFile("Шаблон", 
                                        extension=".xlsx")
        self.template.shouldBeDeleted = False
        self.tkePrev = self.mng.getFile("ТКЕ начало года", 
                                        exactMatch=True, extension=".xlsx")
        self.tkePrev.shouldBeDeleted = False
        self.tkeDK = self.mng.getFile("ТКЕ ДК", 
                                        extension=".xlsx", exactMatch=True)
        self.tkeDK.shouldBeDeleted = False
        self.promDK = self.mng.getFile("Промисловість_ ДК", 
                                        extension=".xlsx", exactMatch=True)
        self.promDK.shouldBeDeleted = False
        self.promPrev = self.mng.getFile("Промисловість_ начало года", 
                                        extension=".xlsx", exactMatch=True)
        self.promPrev.shouldBeDeleted = False
                                        
        
        self.mng.getFile("Оборотно-сальдова вiдомiсть", exactMatch=True)
        self.mng.getFile("Оборотно-сальдова вiдомiсть последний месяц", exactMatch=True)
        self.mng.getFile("ТКЕ", exactMatch=True)
        self.mng.getFile("gpg")
        self.mng.getFile("Промисловість_", exactMatch=True)
        self.mng.getFile("Оборотно-сальдова вiдомiсть пром", exactMatch=True)
        self.mng.getFile("Оборотно-сальдова вiдомiсть ТКЕ", exactMatch=True)

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
                                    extension=".xlsx", exactMatch=True)
            self.saldoTKE = self.mng.getFile("Оборотно-сальдова вiдомiсть ТКЕ", 
                                    extension=".xlsx", exactMatch=True)
            self.saldoProm = self.mng.getFile("Оборотно-сальдова вiдомiсть пром", 
                                    extension=".xlsx", exactMatch=True)
            
            
            if self.mng.getNumberOfFiles() != self.numberOfFilesToStart:
                    raise AttributeError
        except AttributeError:
            print("Не хватает файлов для работы. Проверьте директорию " + str(path))
            msg = r"""Файлы, нужные для работы : 
            1. gpg... : за период 1 января - последняя декада
                        (Менеджер отчетов\Диспетчерський газ\
                            Використання природного газу... (2480bk))
            2. Оборотно-сальдова вiдомiсть : 
                                1 января - по декаду, 
                                категории : 
                                            населення,
                                            населення (газовий депозит), 
                                            бюджет, 
                                            релігійні організації, 
                                            вічний вогонь
                                без лимитов
                        (Менеджер отчетов\Стан розрахунків\
                            Оборотно-сальдовая\
                                Оборотно-сальдова відомість... (2gv))
            3. Оборотно-сальдова вiдомiсть последний месяц : 
                                1 января - 30(31) предыдущий месяц
                                категории : 
                                            населення,
                                            населення (газовий депозит), 
                                            бюджет, 
                                            релігійні організації, 
                                            вічний вогонь
                                без лимитов
                        (Менеджер отчетов\Стан розрахунків\
                            Оборотно-сальдовая\
                                Оборотно-сальдова відомість... (2gv))
            4. Оборотно-сальдова вiдомiсть ТКЕ : 
                                1 января - по декаду, 
                                категории : 
                                            ТЕ,
                                            ТЕ (газовий депозит),
                                            БО, 
                                            КП, 
                                            РО, 
                                            НС, 
                                            ВТЕ
                                без лимитов
                        (Менеджер отчетов\Стан розрахунків\
                            Оборотно-сальдовая\
                                Оборотно-сальдова відомість... (2gv))
            5. Оборотно-сальдова вiдомiсть пром : 
                                1 января - по декаду, 
                                категория промисловість
                                без лимитов
                        (Менеджер отчетов\Стан розрахунків\
                            Оборотно-сальдовая\
                                Оборотно-сальдова відомість... (2gv))
            6. ТКЕ : база Зубарева, 
                    1 января - по декаду
                        (Менеджер отчетов\Стан розрахунків\
                            Розрахунки_ТКЕ_2018-2019)
            7. Промисловість_ : база Зубарева, 
                    1 января - по декаду
                        (Менеджер отчетов\Стан розрахунків\
                            Розрахунки_ПРОМ)

            Файлы, которые не надо загружать, 
            но они должны быть в папке
            1. Шаблон : пустой шаблон
            2. ТКЕ начало года : база Зубарева, 
                                построена на 1.01 текущего года
                        (Менеджер отчетов\Стан розрахунків\
                            Розрахунки_ТКЕ_2018-2019)
            3. Промисловість_ начало года : база Зубарева, 
                                построена на 1.01 текущего года
                        (Менеджер отчетов\Стан розрахунків\
                            Розрахунки_ПРОМ)
            4. ТКЕ ДК : база за 2010-2011 год, 
                                построена в 2018 году
                        (Dept\Monitoring.Gas\4. Газ України\
                            ДК Газ України.  База ТКЕ 2010-2011)
            5. Промисловість_ ДК : база за 2010-2011 год, 
                                построена в 2018 году
                        (Dept\Monitoring.Gas\4. Газ України\
                            ДК Газ України. Промисловість)
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
                self.saldoLastMonth.close()
                self.tke.close()
                self.gasConsumption.close()
                self.prom.close()
                self.saldoTKE.close()
                self.saldoProm.close()
            except:
                print(bcolors.WARNING +\
                    "Программа не смогла закрыть экселевские файлы"\
                    + bcolors.ENDC)
        self.mng.deleteClosedFiles()
        return

    def run(self):
        try:
            self.price = self.getPrice()
            print("Прочитанная цена из файла : ", 
                    self.price, "грн за 1000 м3")
        except (FileNotFoundError, ArithmeticError):
            print("Неправильно указана цена в файле Цена.txt или нет файла")
            print("Заполните файл правильной ценой")
            print("Или же создайте файл (если его нет)")
            print("Исправьте проблему и перезапустите программу")
            self.deleteFiles()
            exit()

        self.naselenie()
        self.religion()
        self.budget()
        self.teploseti()
        self.promishlennost()
        self.forPresident()
        self.generations()
        self.template.save(self.template.pathToFile, 
                            "На печать", extension=".xlsx")
        self.deleteFiles()

        return

    def naselenie(self):
        
        self.template.open(data_only=False)
        templateWsNas = self.template.getWs("Населення")
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()
        self.saldoLastMonth.open(data_only=True)
        saldoLastMonthWs = self.saldoLastMonth.getWs()
        self.gasConsumption.open(data_only=True)
        gasConsumptionWs = self.gasConsumption.getWs("За період")

        self.naselenieIterateInTOVandPAT(templateWsNas, saldoWs, 
                                saldoLastMonthWs, gasConsumptionWs)

        return

    def religion(self):

        self.template.open(data_only=False)
        templateWsReligion = self.template.getWs("Релігія")
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()
        self.saldoLastMonth.open(data_only=True)
        saldoLastMonthWs = self.saldoLastMonth.getWs()
        self.gasConsumption.open(data_only=True)
        gasConsumptionWs = self.gasConsumption.getWs("За період")

        self.religionIterateInTOVandPAT(templateWsReligion, saldoWs, 
                                saldoLastMonthWs, gasConsumptionWs)

        return

    def budget(self):

        self.template.open(data_only=False)
        templateWsBudget = self.template.getWs("Бюджет")
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()
        self.saldoLastMonth.open(data_only=True)
        saldoLastMonthWs = self.saldoLastMonth.getWs()
        self.gasConsumption.open(data_only=True)
        gasConsumptionWs = self.gasConsumption.getWs("За період")

        self.budgetIterateInTOVandPAT(templateWsBudget, saldoWs, 
                                saldoLastMonthWs, gasConsumptionWs)

        return

    def teploseti(self):

        self.template.open(data_only=False)
        templateWsTeploseti = self.template.getWs("Тепломережі")
        self.tkeDK.open(data_only=True)
        tkeDKWs = self.tkeDK.getWs()
        self.tkePrev.open(data_only=True)
        tkePrevWs = self.tkePrev.getWs()
        self.tke.open(data_only=True)
        tkeWs = self.tke.getWs("База")

        # Set range to max line -2 because of merged cell in the bottom
        # of the sheet
        rangeIterInTemplate = "A9" + ":" + "A" +\
                                     str(templateWsTeploseti.max_row-2)
        self.teplosetiIterInRegions(templateWsTeploseti, tkePrevWs, 
                                    tkeWs, tkeDKWs, rangeIterInTemplate)

        return

    def promishlennost(self):
        
        self.template.open(data_only=False)
        templateWsProm = self.template.getWs("Промисловість")
        self.promDK.open(data_only=True)
        promDKWs = self.promDK.getWs()
        self.promPrev.open(data_only=True)
        promPrevWs = self.promPrev.getWs()
        
        # Delete Naftogaz trading data
        self.deleteCompanyData(self.prom, ["42399676"])
        self.prom.open(data_only=True)
        promWs = self.prom.getWs("База_2")
        
        # Set range to max line -2 because of merged cell in the bottom
        # of the sheet
        rangeIterInTemplate = "A9" + ":" + "A" +\
                                        str(templateWsProm.max_row-2)
        self.promIterInRegions(templateWsProm, promDKWs, 
                                promWs, promPrevWs, rangeIterInTemplate)

        return

    def forPresident(self):

        self.template.open(data_only=False)
        templateWsPivot = self.template.getWs("Зведена")
        self.saldo.open(data_only=True)
        saldoNasBudgRelWs = self.saldo.getWs()
        self.saldoTKE.open(data_only=True)
        saldoTKEWs = self.saldoTKE.getWs()
        self.saldoProm.open(data_only=True)
        saldoPromWs = self.saldoProm.getWs()

        numberOfRowsInPAT = self.template.getWs("Населення").max_row
        rangeIterPAT = "B7" + ":" + "B" + str(numberOfRowsInPAT)
        rangeIterTOV = "C7" + ":" + "C" + str(numberOfRowsInPAT)

        naselenie = 0
        budget = 0
        religion = 0
        TKE = 0
        prom = 0
        
        naselenie += self.getTotalPaymentFromSaldo(self.template.getWs(
                                    "Населення"),
                                    rangeIterPAT,
                                    saldoNasBudgRelWs,
                                    [
                                        "населення", 
                                        "населення (газовий депозит)"
                                    ],
                                    "A")
        naselenie += self.getTotalPaymentFromSaldo(self.template.getWs(
                                    "Населення"),
                                    rangeIterTOV,
                                    saldoNasBudgRelWs,
                                    [
                                        "населення", 
                                        "населення (газовий депозит)"
                                    ],
                                    "A")

        budget += self.getTotalPaymentFromSaldo(self.template.getWs(
                                                "Бюджет"),
                                                rangeIterPAT,
                                                saldoNasBudgRelWs,
                                                ["бюджет"],
                                                "A")
        budget += self.getTotalPaymentFromSaldo(self.template.getWs(
                                                "Бюджет"),
                                                rangeIterTOV,
                                                saldoNasBudgRelWs,
                                                ["бюджет"],
                                                "A")

        religion += self.getTotalPaymentFromSaldo(self.template.getWs(
                                    "Релігія"),
                                    rangeIterPAT,
                                    saldoNasBudgRelWs,
                                    [
                                        "релігійні організації", 
                                        "вічний вогонь"
                                    ],
                                    "A")
        religion += self.getTotalPaymentFromSaldo(self.template.getWs(
                                    "Релігія"),
                                    rangeIterTOV,
                                    saldoNasBudgRelWs,
                                    [
                                        "релігійні організації", 
                                        "вічний вогонь"
                                    ],
                                    "A")
        
        column = columnIndexFromString("T")
        TKE = saldoTKEWs.cell(column=column, row=9).value
        
        prom = saldoPromWs.cell(column=column, row=9).value
        naftogazTrading = self.findInSaldo(
                            saldoSheet=saldoPromWs,
                            whatToFind='ТОВ "ГАЗОПОСТАЧАЛЬНА КОМПАНІЯ "НАФТОГАЗ ТРЕЙДИНГ"',
                            whatColumn="T",
                            whatCategory=None,
                            whatResource=None
                            )
        if naftogazTrading != None:
            prom -= naftogazTrading

        columnWhereToWrite = columnIndexFromString("J")
        try:
            templateWsPivot.cell(column=columnWhereToWrite, 
                                row=13).value = naselenie / 1000
            templateWsPivot.cell(column=columnWhereToWrite, 
                                row=14).value = budget / 1000
            templateWsPivot.cell(column=columnWhereToWrite, 
                                row=15).value = TKE / 1000
            templateWsPivot.cell(column=columnWhereToWrite, 
                                row=16).value = prom / 1000
            templateWsPivot.cell(column=columnWhereToWrite, 
                                row=19).value = religion / 1000
        except (UnboundLocalError, AttributeError):
            print("Не заполнен последний столбец во вкладке со сводной таблицей")

        return

    def generations(self):

        self.template.open(data_only=False)
        templateWsPivot = self.template.getWs("Зведена")
        self.promPrev.open(data_only=True)
        promLastWs = self.promPrev.getWs()
        self.prom.open(data_only=True)
        promWs = self.prom.getWs("База_2")

        values = []
        rangeIter = "C11" + ":" + "C" + str(promLastWs.max_row)
        columnIG = columnIndexFromString("IG")

        # Get cell with "Генерації" name in it. This cell is needed 
        # for row number
        cell = self.promPrev.getFirstCellByCriteria("Генерації", 
                                                        rangeIter)
        value = promLastWs.cell(column=columnIG, row=cell.row).value
        values.append(value)

        rangeIter = "H11" + ":" + "H" + str(promLastWs.max_row)
        columnList = ["AB", "CX", "CY", "CZ", "DS", "DQ"]
        columnNumber = []
        for column in columnList:
            columnNumber.append(columnIndexFromString(column))
        # Get cell with "Генерації" name in it. This cell is needed 
        # for row number
        cell = self.prom.getFirstCellByCriteria("Генерації", 
                                                        rangeIter,
                                                        "База_2")
        # Get data form columns "AB", "CX", "CY", "CZ", "DS", "DQ"
        # from Промисловість_
        for column in columnNumber:
            value = promWs.cell(column=column, row=cell.row).value
            values.append(value)
        
        columnList = ["B", "C", "D", "E", "F", "H", "I"]
        columnNumber = []
        for column in columnList:
            columnNumber.append(columnIndexFromString(column))
        if len(values) != 7 or len(values) != len(columnNumber):
            print("Возможны проблемы в заполнении сводной таблицы",
                    "по категории 2.1. Генерації")
        for column, value in zip(columnNumber, values):
            templateWsPivot.cell(column=column, row=17).value = value / 1000
        

    def getTotalPaymentFromSaldo(self, templateSheet, rangeIter, 
                    saldoSheet, whatCategory: list, inWhatColumnFind):
            
        returnValue = 0
        for cells in templateSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    companyName = cell.value
                    value = self.findInSaldo(saldoSheet, companyName, 
                                                whatCategory, None, 
                                                "T", inWhatColumnFind)
                    if value != None:
                        returnValue += value
        return returnValue

    
    def naselenieIterateInTOVandPAT(self, templateSheet, saldoSheet, 
                            saldoLastMonthSheet, gasConsumtionSheet):
        """
        """
        numberOfRows = templateSheet.max_row
        # Iterate in column with TOV companies
        rangeIter = "B7" + ":" + "B" + str(numberOfRows)
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        [
                                            "населення", 
                                            "населення (газовий депозит)"
                                        ], 
                                        None, "G", "D")
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        [
                                            "населення", 
                                            "населення (газовий депозит)"
                                        ], 
                                        ["2019"], "U", "E")
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        [
                                            "населення", 
                                            "населення (газовий депозит)"
                                        ], 
                                        ["2020"], "T", "H")
        self.columnF(templateSheet, saldoLastMonthSheet, gasConsumtionSheet, 
                    rangeIter, "J", [
                                        "населення", 
                                        "населення (газовий депозит)"
                                    ])
        self.columnG(templateSheet, saldoLastMonthSheet, gasConsumtionSheet,
                    rangeIter, "J", [
                                        "населення", 
                                        "населення (газовий депозит)"
                                    ])
        self.columnIandK(templateSheet, rangeIter)
        self.columnJ(templateSheet, saldoSheet, rangeIter, 
                                    [
                                        "населення", 
                                        "населення (газовий депозит)"
                                    ])
        
        # Iterate in column with PAT companies
        rangeIter = "C7" + ":" + "C" + str(numberOfRows)
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        [
                                            "населення", 
                                            "населення (газовий депозит)"
                                        ], 
                                        None, "G", "D")
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        [
                                            "населення", 
                                            "населення (газовий депозит)"
                                        ], 
                                        ["2019"], "U", "E") 
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        [
                                            "населення", 
                                            "населення (газовий депозит)"
                                        ], 
                                        ["2020"], "T", "H")
        self.columnF(templateSheet, saldoLastMonthSheet, gasConsumtionSheet, 
                    rangeIter, "J", [
                                        "населення", 
                                        "населення (газовий депозит)"
                                    ])
        self.columnG(templateSheet, saldoLastMonthSheet, gasConsumtionSheet,
                    rangeIter, "J", [
                                        "населення", 
                                        "населення (газовий депозит)"
                                    ])
        self.columnIandK(templateSheet, rangeIter)
        self.columnJ(templateSheet, saldoSheet, rangeIter, 
                                    [
                                        "населення", 
                                        "населення (газовий депозит)"
                                    ])

        return

    def religionIterateInTOVandPAT(self, templateSheet, saldoSheet, 
                                saldoLastMonthSheet, gasConsumtionSheet):
        """
        """
        numberOfRows = templateSheet.max_row
        # Iterate in column with TOV companies
        rangeIter = "B7" + ":" + "B" + str(numberOfRows)
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        [
                                            "релігійні організації", 
                                            "вічний вогонь"
                                        ], 
                                        None, "G", "D")
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        [
                                            "релігійні організації", 
                                            "вічний вогонь"
                                        ], 
                                        ["2019"], "U", "E")
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        [
                                            "релігійні організації", 
                                            "вічний вогонь"
                                        ], 
                                        ["2020"], "T", "H")
        self.columnF(templateSheet, saldoLastMonthSheet, gasConsumtionSheet, 
                        rangeIter, "W", [
                                            "релігійні організації", 
                                            "вічний вогонь"
                                        ])
        self.columnG(templateSheet, saldoLastMonthSheet, gasConsumtionSheet,
                        rangeIter, "W", [
                                            "релігійні організації", 
                                            "вічний вогонь"
                                        ])
        self.columnIandK(templateSheet, rangeIter)
        self.columnJ(templateSheet, saldoSheet, 
                        rangeIter, [
                                        "релігійні організації", 
                                        "вічний вогонь"
                                    ])
        
        # Iterate in column with PAT companies
        rangeIter = "C7" + ":" + "C" + str(numberOfRows)
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        [
                                            "релігійні організації", 
                                            "вічний вогонь"
                                        ], 
                                        None, "G", "D")
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        [
                                            "релігійні організації", 
                                            "вічний вогонь"
                                        ], 
                                        ["2019"], "U", "E") 
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        [
                                            "релігійні організації", 
                                            "вічний вогонь"
                                        ], 
                                        ["2020"], "T", "H")
        self.columnF(templateSheet, saldoLastMonthSheet, gasConsumtionSheet, 
                        rangeIter, "W", [
                                            "релігійні організації", 
                                            "вічний вогонь"
                                        ])
        self.columnG(templateSheet, saldoLastMonthSheet, gasConsumtionSheet,
                        rangeIter, "W", [
                                            "релігійні організації", 
                                            "вічний вогонь"
                                        ])
        self.columnIandK(templateSheet, rangeIter)
        self.columnJ(templateSheet, saldoSheet, 
                        rangeIter, [
                                        "релігійні організації", 
                                        "вічний вогонь"
                                    ])

        return

    def budgetIterateInTOVandPAT(self, templateSheet, saldoSheet, 
                            saldoLastMonthSheet, gasConsumtionSheet):
        """
        """
        numberOfRows = templateSheet.max_row
        # Iterate in column with TOV companies
        rangeIter = "B7" + ":" + "B" + str(numberOfRows)
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        ["бюджет"], None, "G", "D")
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        ["бюджет"], ["2019"], "U", "E")
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        ["бюджет"], ["2020"], "T", "H")
        self.columnF(templateSheet, saldoLastMonthSheet, gasConsumtionSheet, 
                    rangeIter, None, ["бюджет"])
        self.columnG(templateSheet, saldoLastMonthSheet, gasConsumtionSheet,
                    rangeIter, None, ["бюджет"])
        self.columnIandK(templateSheet, rangeIter)
        self.columnJ(templateSheet, saldoSheet, rangeIter, ["бюджет"])
        
        # Iterate in column with PAT companies
        rangeIter = "C7" + ":" + "C" + str(numberOfRows)
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        ["бюджет"], None, "G", "D")
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        ["бюджет"], ["2019"], "U", "E") 
        self.findInSaldoWriteToTemplate(templateSheet, saldoSheet, rangeIter,
                                        ["бюджет"], ["2020"], "T", "H")
        self.columnF(templateSheet, saldoLastMonthSheet, gasConsumtionSheet, 
                    rangeIter, None, ["бюджет"])
        self.columnG(templateSheet, saldoLastMonthSheet, gasConsumtionSheet,
                    rangeIter, None, ["бюджет"])
        self.columnIandK(templateSheet, rangeIter)
        self.columnJ(templateSheet, saldoSheet, rangeIter, ["бюджет"])

        return
        
    def findInSaldoWriteToTemplate(self, templateSheet, saldoSheet, rangeIter, 
                                    whatCategory: list, whatResource: list, 
                                    whatColumn: str, whereToPut: str):
        """
        """
        columnWhereToPut = columnIndexFromString(whereToPut)
        for cells in templateSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    companyName = cell.value
                    value = self.findInSaldo(saldoSheet, 
                                            whatToFind=companyName,
                                            whatCategory=whatCategory,
                                            whatResource=whatResource,
                                            whatColumn=whatColumn)
                    if value != None:
                        templateSheet.cell(column=columnWhereToPut, 
                                        row=cell.row).value = value / 1000
        return

    def findInSaldo(self, saldoSheet, whatToFind: str, 
                    whatCategory: list, whatResource: list, whatColumn: str, inWhatColumnFind="A"):
        """
        """
        numberOfRows = saldoSheet.max_row
        rangeIter = inWhatColumnFind + "10" + ":" + inWhatColumnFind + str(numberOfRows)
        columnCategory = columnIndexFromString("C")
        columnResource = columnIndexFromString("F")
        columnWithData = columnIndexFromString(whatColumn)
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
                                                row=row).value).strip()
                        value = saldoSheet.cell(column=columnWithData,
                                                row=row).value
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
                    
                    return returnValue
        
        # If company wasn`t found in saldo then set return value to 0
        try:
            returnValue
        except UnboundLocalError:
            returnValue = 0
            return returnValue
        
        return
        
    def columnF(self, templateSheet, saldoSheet, gasConsumptionSheet, 
                rangeIter: str, whatColumnInConsumption: str, whatCategory: list):
        """
        """
        columnWhereToPut = columnIndexFromString("F")
        for cells in templateSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    companyName = cell.value

                    # Working with F column in template sheet
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
                    templateSheet.cell(column=columnWhereToPut, row=cell.row).value = summary / 1000
        return

    def columnG(self, templateSheet, saldoSheet, gasConsumptionSheet,
                rangeIter: str, whatColumnInConsumption: str, whatCategory: list):
        """
        """
        columnWhereToPut = columnIndexFromString("G")
        for cells in templateSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    companyName = cell.value

                    # Working with G column in template sheet
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

                    summary = valueFromGasConsumption / 1000 *\
                            self.price + valueFromSaldoLastMonth / 1000 
                    templateSheet.cell(column=columnWhereToPut, row=cell.row).value = summary
        return


    def findInGasConsumption(self, gasConsumptionSheet, whatToFind: str, whatColumn: str):
        """
        """
        columnWithData = columnIndexFromString(whatColumn)
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

    def columnIandK(self, templateSheet, rangeIter):
        """
        """
        columnWithPercent = columnIndexFromString("I")
        columnWithDebt = columnIndexFromString("K")
        columnPaymentForConsumedGas = columnIndexFromString("H")
        columnAmountConsumedGas = columnIndexFromString("G")
        for cells in templateSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    # Find in Декадка numbers in columns H, G with value that were divided by 1000
                    # thats why they should be multiplied by 1000
                    paymentForConsumedGas = templateSheet.cell(column=columnPaymentForConsumedGas, 
                                            row=cell.row).value * 1000
                    amountConsumedGas = templateSheet.cell(column=columnAmountConsumedGas, 
                                            row=cell.row).value * 1000
                    
                    # Fill column I in template
                    if amountConsumedGas == 0:
                        templateSheet.cell(column=columnWithPercent, 
                                            row=cell.row).value = 0
                    elif amountConsumedGas != 0:
                        templateSheet.cell(column=columnWithPercent, 
                                            row=cell.row).value = \
                                            paymentForConsumedGas / amountConsumedGas * 100

                    # Fill column K in template
                    templateSheet.cell(column=columnWithDebt, 
                                            row=cell.row).value = (amountConsumedGas - paymentForConsumedGas) / 1000
        return

    def columnJ(self, templateSheet, saldoSheet, 
                        rangeIter: str, whatCategory: list):
        """
        """
        columnWithDebtPreviousYears = columnIndexFromString("D")
        columnConsumedGasAmount = columnIndexFromString("G")
        columnPaymentForConsumedGas = columnIndexFromString("H")
        columnWhereToPut = columnIndexFromString("J")
        for cells in templateSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    # Find in Декадка numbers in columns D, G, H with value that were divided by 1000
                    # thats why they should be multiplied by 1000
                    lastPeriodDebt = templateSheet.cell(column=columnWithDebtPreviousYears,
                                            row=cell.row).value * 1000
                    consumedGasAmount = templateSheet.cell(column=columnConsumedGasAmount,
                                            row=cell.row).value * 1000
                    paymentForConsumedGas = templateSheet.cell(column=columnPaymentForConsumedGas,
                                            row=cell.row).value * 1000

                    valueFromColumnTSaldo = self.findInSaldo(saldoSheet, 
                                            whatToFind=cell.value, 
                                            whatCategory=whatCategory,
                                            whatResource=None,
                                            whatColumn="T")
                    valueFromColumnTSaldoThisYear = self.findInSaldo(saldoSheet, 
                                            whatToFind=cell.value, 
                                            whatCategory=whatCategory,
                                            whatResource=["2020"],
                                            whatColumn="T")
                    valueFromColumnJSaldo = self.findInSaldo(saldoSheet, 
                                            whatToFind=cell.value, 
                                            whatCategory=whatCategory,
                                            whatResource=None,
                                            whatColumn="J")
                    
                    total = lastPeriodDebt - valueFromColumnTSaldo + \
                                valueFromColumnTSaldoThisYear + valueFromColumnJSaldo + \
                                consumedGasAmount - paymentForConsumedGas
                    
                    templateSheet.cell(column=columnWhereToPut, row=cell.row).value = total / 1000
        return

    def teplosetiIterInRegions(self, templateSheet, tkePrevWs, 
                                tkeWs, tkeDKsheet, rangeIter):
        """
        """
        for cells in templateSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    region = cell.value

                    listOfTwoValues = self.teplosetiColumnB(tkePrevWs, 
                                                    tkeDKsheet, region)
                    # Get debt of NAK and tkeDK together
                    summOfTwoValues = listOfTwoValues[0] + listOfTwoValues[1]
                    templateSheet.cell(column=cell.column+1, 
                            row=cell.row).value = summOfTwoValues / 1000
                                    
                    listOfValues = self.teplosetiColumnsFromCToF(tkeWs,
                                                        region, 
                                                        ["DI", "DJ",
                                                        "DK", "DQ",
                                                        "CU", "CG",
                                                        "GR", "GS"])
                    valueDI = listOfValues[0]
                    templateSheet.cell(column=cell.column+2, 
                            row=cell.row).value = valueDI / 1000
                    valueGR = listOfValues[6] # gas deposit                
                    valueDJ = listOfValues[1]
                    templateSheet.cell(column=cell.column+3, 
                            row=cell.row).value = (valueDJ+valueGR) / 1000
                    
                    valueGS = listOfValues[7] # gas deposit    
                    valueDK = listOfValues[2]
                    templateSheet.cell(column=cell.column+4, 
                            row=cell.row).value = (valueDK+valueGS) / 1000

                    valueDQ = listOfValues[3]
                    templateSheet.cell(column=cell.column+5, 
                            row=cell.row).value = (valueDQ+valueGS) / 1000
                                    
                    valueCU = listOfValues[4]
                    valueCG = listOfValues[5]
                    valueFromListOfTwoValues = listOfTwoValues[1]
                    total = valueDI + valueCU + valueCG + \
                            valueFromListOfTwoValues + valueDK - valueDQ

                    templateSheet.cell(column=cell.column+7, 
                            row=cell.row).value = total / 1000
                    
                    if valueDK != 0:
                        templateSheet.cell(column=cell.column+6, 
                            row=cell.row).value = (valueDQ+valueGS) / (valueDK+valueGS) * 100
                    elif valueDK == 0:
                        templateSheet.cell(column=cell.column+6, 
                            row=cell.row).value = 0

                    templateSheet.cell(column=cell.column+8, 
                            row=cell.row).value = (valueDK-valueDQ) / 1000


                    
    def teplosetiColumnB(self, tkeSheet, tkeDKsheet, regionName):
        """
        """
        rangeIterInTkePrev = "B15" + ":" + "B" + str(tkeSheet.max_row)
        columnDebtInTke = columnIndexFromString("EO")
        returnValueFromTke = 0
        isDone = False
        for cells in tkeSheet[rangeIterInTkePrev]:
            for cell in cells:
                # If cell in P column is equal to ВАТ/ПАТ
                if tkeSheet.cell(column=columnIndexFromString("P"), 
                                    row=cell.row).value == "РАЗОМ:":
                    isDone = True
                    break
                
                if cell.value == regionName:
                    returnValueFromTke += tkeSheet.cell(column=columnDebtInTke, row=cell.row).value
            if isDone:
                break

        rangeIterInTkeDK = "B12" + ":" + "B" + str(tkeDKsheet.max_row)
        columnDebtInTkeDK = columnIndexFromString("EG")
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

    def teplosetiColumnsFromCToF(self, tkeSheet, 
                            regionName, whatColumns: list):
        """
        """
        listOfColumns = [columnIndexFromString(column) for column in whatColumns]
        rangeIterInTke = "C15" + ":" + "C" + str(tkeSheet.max_row)
        returnValuesList = [0]*len(listOfColumns)
        isDone = False

        for cells in tkeSheet[rangeIterInTke]:
            for cell in cells:
                # If null then its end of data
                if cell.value == "Область / Населений пункт":
                    isDone = True
                    break
                if cell.value == regionName:
                    for i in range(0, len(listOfColumns)):
                        returnValuesList[i] += tkeSheet.cell(column=listOfColumns[i], 
                                                                row=cell.row).value        
            if isDone:
                break
            
        for value in returnValuesList:
            if value == None:
                value = 0
        
        return returnValuesList

    def promColumnB(self, promPrevSheet, promDkSheet, regionName):
        """
        """
        rangeIterInPromPrev = "D12" + ":" + "D" + str(promPrevSheet.max_row)
        columnTotalDebt = columnIndexFromString("IG")
        returnValuePromPrev = 0
        isDone = False
        for cells in promPrevSheet[rangeIterInPromPrev]:
            for cell in cells:
                # If cell equal to "Всього:" then it is the end of the sheet
                row = cell.row
                if promPrevSheet.cell(column=columnIndexFromString("C"),
                                        row=row).value == "Всього:":
                    isDone = True
                    break
                # Add debt of all regions in sheet
                if cell.value == regionName:
                    # Check if its region in total with all
                    # companies but not company itself in that region
                    if promPrevSheet.cell(column=columnIndexFromString("C"),
                                            row=row).value == None:
                        break

                    debt = promPrevSheet.cell(column=columnTotalDebt, 
                                                row=cell.row).value
                    # Check for not None type
                    if debt == None:
                        debt = 0
                    returnValuePromPrev += debt
            
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
                    debt = promDkSheet.cell(column=columnTotalDebt, 
                                                row=cell.row).value
                    if debt == None:
                        debt = 0
                    returnValuePromDK += debt
            
            if isDone:
                break
        
        return [returnValuePromPrev, returnValuePromDK]

    def promColumnsFromCToF(self, promSheet, region, fromWhatColumn: str):
        """
        """
        rangeIterInProm = "H12" + ":" + "H" + str(promSheet.max_row)
        columnWithData = columnIndexFromString(fromWhatColumn)
        returnValue = 0
        for cells in promSheet[rangeIterInProm]:
            for cell in cells:                
                if cell.value == region:
                    returnValue += promSheet.cell(column=columnWithData, 
                                                    row=cell.row).value
        
        if returnValue != None:
            return returnValue
        elif returnValue == None:
            return 0

    def deleteCompanyData(self, prom, listOfSpecificCompanies):
        """
        """
        # Save to temp variable prom excel book
        tempProm = self.prom
        # Open with pyWin
        self.prom = self.mng.addFileByPath(
                                            self.prom.pathToFile, 
                                            self.prom.fileName,
                                            returnFile=True, 
                                            defaultParser=False, 
                                            openBy=1
                                            )
        self.prom.open()
        promWs = self.prom.getWs("База_2")

        column = columnIndexFromString("E")
        maxColumn = openpyxl.utils.get_column_letter(promWs.UsedRange.Columns.Count)

        for row in range(12, promWs.UsedRange.Rows.Count):
            EDRPOU = promWs.Cells(row, column).Value
            if str(EDRPOU) in listOfSpecificCompanies:
                promWs.Range("A"+str(row)+":"+maxColumn+str(row)).Clear()
        
        # Save file with rewriting
        self.prom.save(self.prom.pathToFile, self.prom.fileNameWithoutExtension, conflictResolution=True)
        self.prom.close()
        # This is needed because two copies of Промисловість_.xlsx were created
        # and one of them should be removed from mng before deleting 
        # because of error
        self.mng.removeFile(self.prom)
        self.prom = tempProm

        return

    def promIterInRegions(self, templateSheet, promDkSheet, 
                            promSheet, promPrevSheet, rangeIter):
        """
        """
        for cells in templateSheet[rangeIter]:
            for cell in cells:
                if cell.value != None:
                    region = cell.value

                    listOfTwoValues = self.promColumnB(promPrevSheet, promDkSheet, region)
                    # Get debt of NAK and DK together
                    summOfTwoValues = listOfTwoValues[0] + listOfTwoValues[1]
                    value1 = listOfTwoValues[1]
                    templateSheet.cell(column=cell.column+1, 
                            row=cell.row).value = summOfTwoValues / 1000

                    value2 = self.promColumnsFromCToF(promSheet, region, "DP")
                    templateSheet.cell(column=cell.column+2, 
                            row=cell.row).value = value2 / 1000
                    
                    templateSheet.cell(column=cell.column+3, row=cell.row).value =\
                            self.promColumnsFromCToF(promSheet, region, "CX") / 1000
                    
                    value3 = self.promColumnsFromCToF(promSheet, region, "CY")
                    templateSheet.cell(column=cell.column+4, 
                            row=cell.row).value = value3 / 1000
                    
                    value4 = self.promColumnsFromCToF(promSheet, region, "CZ")
                    templateSheet.cell(column=cell.column+5, 
                            row=cell.row).value = value4 / 1000

                    if value3 != 0:
                        templateSheet.cell(column=cell.column+6, 
                            row=cell.row).value = value4 / value3 * 100
                    elif value3 == 0:
                        templateSheet.cell(column=cell.column+6, 
                            row=cell.row).value = 0

                    value5 = self.promColumnsFromCToF(promSheet, region, "AB")
                    value6 = self.promColumnsFromCToF(promSheet, region, "DO")
                    total = value5 + value6 + value2 + value1 + value3 - value4
                    templateSheet.cell(column=cell.column+7, 
                                row=cell.row).value = total / 1000
                    templateSheet.cell(column=cell.column+8, 
                                row=cell.row).value = (value3 - value4) / 1000

    def getPrice(self):
        """Get price value from Цена.txt
        """
        with open(self.template.pathToFile+"\\"+"Цена.txt", "r") as f:
            content = f.read().splitlines()
            while True:
                haveEmpty = False
                for element in content:
                    if element == "":
                        haveEmpty = True
                        content.remove("")
                if haveEmpty == False:
                    break
        if not content:
            print(bcolors.WARNING +\
                """Файл "Цена.txt" пустой"""\
                + bcolors.ENDC)
            price = 0
            raise ArithmeticError
        else:
            price = float(content[0])
        return price