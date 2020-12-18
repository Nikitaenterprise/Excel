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
                                    1.01.11 - 30 предыдущего месяца, 
                                    категория :
                                                ВТВ+НОРМ,
                                                населення
                                    признаки субъекта Усі ВАТ+ЗБУТ
                                    без лимитов
                                    без судов
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
        rangeIterNas = "B4" + ":" + "B" + str(templateWsNasel.max_row)
        columnWithDebt = openpyxl.utils.column_index_from_string("C")
        summary = []
        for cells in templateWsNasel[rangeIterNas]:
            for cell in cells:
                if (cell.value != None and
                    cell.value != "Всього:"):
                    company = cell.value
                    data = findInSaldo(saldoWs,
                                        company,
                                        ["населення"],
                                        ["2020", "2021"],
                                        ["I", "T", "U"])
                    summary.append(data)

                    if data[0] == 0:
                        percent = 0
                    elif data[0] != 0:
                        percent = data[1]/data[0] * 100
                    self.write(templateWsNasel, 
                                columnWithDebt,
                                cell.row,
                                data[2],
                                percent)
                
                elif cell.value == "Всього:":
                    data0, data1, data2 = 0, 0, 0
                    for i in summary:
                        data0 += i[0]
                        data1 += i[1]
                        data2 += i[2]

                    if data0 == 0:
                        percent = 0
                    elif data0 != 0:
                        percent = data1/data0 * 100
                    self.write(templateWsNasel, 
                                columnWithDebt,
                                cell.row,
                                data2,
                                percent)

        
        templateWsVTV = self.template.getWs("ВТВ+НОРМ")
        rangeIterVTV = "B4" + ":" + "B" + str(templateWsVTV.max_row)
        summary = []
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
                    summary.append(data)
                    
                    if data[3] == 0:
                        percent = 0
                    elif data[3] != 0:
                        percent = (data[1] - data[2])/data[3] * 100
                    self.write(templateWsVTV, 
                                columnWithDebt,
                                cell.row,
                                data[0],
                                percent)
                
                if cell.value == "Всього:":
                    data0, data1, data2, data3 = 0, 0, 0, 0
                    for i in summary:
                        data0 += i[0]
                        data1 += i[1]
                        data2 += i[2]
                        data3 += i[3]
                        # Check for excluding negative percents
                        if (data1 - data2) < 0:
                            data3 = 0

                    if data3 == 0:
                        percent = 0
                    elif data3 != 0:
                        percent = (data1 - data2)/data3 * 100
                    self.write(templateWsVTV, 
                                columnWithDebt,
                                cell.row,
                                data0,
                                percent)
                    

        self.template.save(self.template.pathToFile, 
                            "Население и ВТВ+НОРМ",
                            extension=".xlsx")
        self.deleteFiles()

    def write(self, ws, column, row, debt, percent):
        
        ws.cell(column=column, row=row).value = debt
        ws.cell(column=column+1, row=row).value = percent
