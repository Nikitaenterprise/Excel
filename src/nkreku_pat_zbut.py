from src.alg import *
from src.saldo import *


class NKREKU_PAT_ZBUT_VTV_naselenie(Algorithm):

    def checkIfDirectoryIsReady(self, path: str):
        self.mng.addFilesInDir()

        self.template = self.mng.getFile("Шаблон", 
                                        extension=".xlsx")
        self.mng.getFile("Оборотно-сальдова вiдомiсть")

        self.mng.deleteUnCalledFiles()
        self.mng.allFromXlsToXlsx()

        try:
            self.saldo = self.mng.getFile("Оборотно-сальдова вiдомiсть",
                                            extension=".xlsx")
            if self.mng.getNumberOfFiles() != self.numberOfFilesToStart:
                raise AttributeError
        except AttributeError:
            print("Не хватает файлов для работы. Проверьте директорию " + str(path))
            msg = r"""Файлы, нужные для работы: 
            1. Оборотно-сальдова вiдомiсть : 
                                    1.01.14 - 30 предыдущего месяца, 
                                    категория :
                                                ВТВ+НОРМ,
                                                населення
                                    признаки субъекта Усі ВАТ+ЗБУТ
                    (Менеджер отчетов\Стан розрахунків\
                        Оборотно-сальдовая\
                            Оборотно-сальдова відомість... (2gv))
            Файлы, которые не надо загружать, но они должны быть в папке
            1. Шаблон : пустой шаблон
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
            except:
                print(bcolors.WARNING +\
                    "Программа не смогла закрыть экселевские файлы"\
                    + bcolors.ENDC)
        self.mng.deleteClosedFiles()
        return

    def run(self):
        self.saldo.open(data_only=True)
        saldoWs = self.saldo.getWs()

        self.template.open(data_only=False)
        templateWsNasel = self.template.getWs("Населення")
        rangeIterNas = "B4" + ":" + "B" + str(templateWsNasel.max_row-1)
        columnWithDebt = openpyxl.utils.column_index_from_string("C")
        for cells in templateWsNasel[rangeIterNas]:
            for cell in cells:
                if (cell.value != None and
                    cell.value != "Всього:" and
                    cell.value != "Разом:"):
                    company = cell.value
                    data = findInSaldo(saldoWs,
                                        company,
                                        ["населення"],
                                        ["2018", "2019"],
                                        ["I", "T", "U"])
                    templateWsNasel.cell(column=columnWithDebt, 
                                        row=cell.row).value = data[2]
                    if data[0] == 0:
                        percent = 0
                    elif data[0] != 0:
                        percent = data[1]/data[0] * 100

                    templateWsNasel.cell(column=columnWithDebt+1, 
                                        row=cell.row).value = percent
        
        templateWsVTV = self.template.getWs("ВТВ+НОРМ")
        rangeIterVTV = "B4" + ":" + "B" + str(templateWsVTV.max_row-1)
        for cells in templateWsVTV[rangeIterVTV]:
            for cell in cells:
                if (cell.value != None and
                    cell.value != "Всього:" and
                    cell.value != "Разом:"):
                    company = cell.value
                    data = findInSaldo(saldoWs,
                                        company,
                                        ["ВТВ + НОРМ"],
                                        None,
                                        ["U", "T", "G", "I"])
                    templateWsVTV.cell(column=columnWithDebt, 
                                        row=cell.row).value = data[0]
                    if data[3] == 0:
                        percent = 0
                    elif data[3] != 0:
                        percent = (data[1] - data[2])/data[3] * 100

                    templateWsVTV.cell(column=columnWithDebt+1, 
                                        row=cell.row).value = percent
                    

        self.template.save(self.template.pathToFile, 
                            "Население и ВТВ+НОРМ",
                            extension=".xlsx")
        self.deleteFiles()