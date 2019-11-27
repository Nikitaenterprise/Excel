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
            1. Шаблон... файл-форма отчетности НКРЭКУ №2 с заполненными 
                        данными в колонке L в строках 48(4.1)-51(4.4)
                    даные заполняются из : 
                    Финансы\Движение денежных средств\Импорт платежей
            2. Оборотно-сальдова вiдомiсть : 
                    за предыдущий месяц (1 - 30),
                    по категориям 
                            ТЕ, БО, КП, РО, НС, ВТЕ,
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
        self.deleteFiles()

    def fill(self):

        self.template.open(data_only=True)
        self.templateWs = self.template.getWs()
        self.saldo.open(data_only=True)
        self.saldoWs = self.saldo.getWs()
        self.saldoEE.open(data_only=True)
        self.saldoEEWs = self.saldoEE.getWs()

        