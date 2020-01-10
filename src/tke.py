from src.alg import *

class TKE(Algorithm):

    def checkIfDirectoryIsReady(self, path: str):

        self.mng.addFilesInDir()

        # Checks if the files are present
        self.mng.getFile("Новый отчет")
        self.mng.getFile("Киiвтеплоенерго")
        self.mng.getFile("Оператор ГТС")
        self.mng.getFile("Звiт_Рестр")
        self.mng.getFile("90%ТКЕ_ПСО")
        # Deletes other files in dir
        self.mng.deleteUnCalledFiles()               
        self.mng.allFromXlsToXlsx()
        
        try:
            self.todayTKE = self.mng.getFile("Новый отчет", extension=".xlsx")
            self.yesterdayTKE = self.mng.getFile("90%ТКЕ_ПСО", extension=".xlsx")
            self.kyivEnergoPas = self.mng.getFile("Киiвтеплоенерго", extension=".xlsx")
            self.operatorGtsPas = self.mng.getFile("Оператор ГТС", extension=".xlsx")
            self.restructurization1730 = self.mng.getFile("Звiт_Рестр", extension=".xlsx")
            
            if self.mng.getNumberOfFiles() != self.numberOfFilesToStart:
                    raise AttributeError
        except AttributeError:
            print("Не хватает файлов для работы. Проверьте директорию " + str(path))
            msg = """Файлы, нужные для работы: 
            1. Новый отчет  свежая база ТКЕ_ПСО
            2. 90%ТКЕ_ПСО_... вчрашняя база ТКЕ_ПСО (там где ... там стоит дата с месяцем). 
                            Это файл, с которым будет сравниваться список предприятий (новички).
            3. Паспорт Киiвтеплоенерго КП ВО Киiвради (КМДА) за сегодня
            4. Звiт_Рестр_1730_Друк_ВсiОбластi  файл по 1730
            Итого: 4 экселевских файла
            После исправления запустите программу заново. Сейчас программа завершит работу
            Нажмите любую клавишу а затем Enter
            """
            print(bcolors.OKGREEN + msg + bcolors.ENDC)
            input()
            exit()

    def deleteFiles(self, programmIsDone=True):
        """Deletes all created files with .xlsx extension
        """
        # If programm has daone its work then close files
        if programmIsDone == True:
            try:
                self.yesterdayTKE.close()
                self.kyivEnergoPas.close()
                self.restructurization1730.close()
            except:
                print(bcolors.WARNING +\
                    "Программа не смогла закрыть экселевские файлы"\
                    + bcolors.ENDC)
        self.mng.addFileByPath(self.todayTKE.pathToFile, 
                            self.generateName() + ".xlsx")
        self.mng.deleteClosedFiles()
        return

    def run(self):
        self.copyColumn()
        self.mainCalculations()
        name = self.generateName()
        self.hideColumns()
        #self.addFilter()
        self.todayTKE.save(self.todayTKE.pathToFile, name, extension=".xls")
        self.deleteFiles()
        
    def mainCalculations(self):

        fileNameForRead = "forRead"
        self.mng.createDuplicate(self.todayTKE, fileNameForRead)
        #Read from/write to block
        # Write to
        self.todayTKE.open(data_only=False)
        todayWs = self.todayTKE.getWs("Sheet1")

        # Read from
        fileNameForRead += ".xlsx"
        todayTkeWithData = self.mng.addFileByPath(self.todayTKE.pathToFile, 
                            fileNameForRead, returnFile=True)
        todayTkeWithData.open(data_only=True)
        todayWsData = todayTkeWithData.getWs("Sheet1")

        # Set 1 and 'договир э' to those companies who have a restructurization contract
        self.restructurization1730.open()
        self.checkRestructurizationData(todayWs, todayWsData)

        numberOfRows = todayWs.max_row
        rangeIter = "O12" + ":" + "O" + str(numberOfRows)
        for cells in todayWsData[rangeIter]:
            for cell in cells:
                if cell.value != "" and cell.value != None:
                    columnWithConditions = columnIndexFromString("AT")
                    # This check needs for empty cell not to be filled
                    if todayWsData.cell(column=columnWithConditions, row=cell.row).value != "":
                        todayWs.cell(column=columnWithConditions, row=cell.row).value = 1
                        todayWsData.cell(column=columnWithConditions, row=cell.row).value = 1

                        # Call restructurization() for get debt
                        summary = self.restructurization(todayWsData, cell.column, cell.row)

                        column = columnIndexFromString("AM")
                        if summary != None:
                            # If company have debt >0 then this summ will appear in
                            # column "AM"
                            todayWs.cell(column=column, row=cell.row).value = summary
                            todayWsData.cell(column=column, row=cell.row).value = summary
                            # Set 0 to the column with conditions
                            todayWs.cell(column=columnWithConditions, row=cell.row).value = 0
                            todayWsData.cell(column=columnWithConditions, row=cell.row).value = 0

                        elif summary == None:
                            # If company have debt or debt <0 then "договір є" will
                            # appear in column "AM"
                            todayWs.cell(column=column, row=cell.row).value = str("договір є")                          
                            todayWsData.cell(column=column, row=cell.row).value = str("договір є")                          
        
        # Set 1 to those companies who are currently controlled by Фонд Держ Майна
        # If they have >60% then set 1
        listOfDerzhMaino = [
                            "00130820",
                            "00130850",
                            "40885849",
                            "00131050",
                            "30083966",
                            "05471158",
                            "00131771"
                            ]
        # Iterate in EDRPOU column
        rangeIter = "P12" + ":" + "P" + str(numberOfRows)
        columnWithPercents = columnIndexFromString("AG")
        columnWithConditions = columnIndexFromString("AT")
        columnBV = columnIndexFromString("BV")
        for cells in todayWsData[rangeIter]:
            for cell in cells:
                if cell.value in listOfDerzhMaino:
                    # Set 60 to BV column for right calculation of debt
                    todayWs.cell(column=columnBV, row=cell.row).value = 60
                    todayWsData.cell(column=columnBV, row=cell.row).value = 60
                    if todayWsData.cell(column=columnWithPercents, row=cell.row).value >= 60:
                        todayWs.cell(column=columnWithConditions, row=cell.row).value = 1
                        todayWsData.cell(column=columnWithConditions, row=cell.row).value = 1
                        # Set to 0 sum to 90%
                        todayWs.cell(column=columnWithPercents + 1, row=cell.row).value = 0                     
                        todayWsData.cell(column=columnWithPercents + 1, row=cell.row).value = 0

        # Copy data from 'поточний лимит' to 'попередний лимит'
        rangeIter1 = "BO10" + ":" + "BU" + str(numberOfRows)
        rangeIter2 = "BB10" + ":" + "BH" + str(numberOfRows)
        for cells1, cells2 in zip(todayWsData[rangeIter1], todayWsData[rangeIter2]):
            for cell1, cell2 in zip(cells1, cells2):
                if cell1.row == cell2.row:
                    todayWs.cell(column=cell2.column, row=cell2.row).value = cell1.value
        
        rangeAG = "AG10" + ":" + "AG" + str(numberOfRows)
        rangeAI = "AI10" + ":" + "AI" + str(numberOfRows)
        rangeAM = "AM10" + ":" + "AM" + str(numberOfRows)
        columnR = columnIndexFromString("R")
        for cells in todayWsData[rangeAG]:
            for cellWithPercent in cells:
                try:
                    # If percents are in between 89 and 90
                    if (cellWithPercent.value > 89 and 
                        cellWithPercent.value < 90):

                        cellWithDebt = todayWsData.cell(
                                    column=cellWithPercent.column + 1, 
                                    row=cellWithPercent.row).value
                        cellWithCompany = todayWsData.cell(
                                    column=columnR, 
                                    row=cellWithPercent.row).value
                        # If debt less than 1
                        if (cellWithDebt != None and 
                            cellWithDebt > 0 and 
                            cellWithDebt < 1 and
                            cellWithCompany != None and
                            cellWithCompany != ""):

                            todayWsData.cell(
                                    column=cellWithPercent.column, 
                                    row=cellWithPercent.row).value = 90
                            todayWs.cell(
                                    column=cellWithPercent.column, 
                                    row=cellWithPercent.row).value = 90
                            todayWs.cell(
                                    column=columnWithConditions, 
                                    row=cellWithPercent.row).value = 1
                            todayWsData.cell(
                                    column=columnWithConditions, 
                                    row=cellWithPercent.row).value = 1
                    
                    # Set 1 to conditions column if payment >90%
                    # this need to be done because of situation when
                    # company has restructurization contract with debt in it
                    cellWithConditions = todayWsData.cell(
                                            column=columnWithConditions, 
                                            row=cellWithPercent.row).value
                    if (cellWithPercent.value > 90 and 
                        cellWithConditions != None and
                        cellWithConditions != ""):
                        
                        todayWs.cell(column=columnWithConditions, 
                                    row=cellWithPercent.row).value = 1
                        todayWsData.cell(column=columnWithConditions, 
                                    row=cellWithPercent.row).value = 1
                except TypeError:
                    continue

        # Check for numbers like 0,001, and if so clear value
        for cells in todayWsData[rangeAI]:
            for cell in cells:
                try:
                    if cell.value > 0 and cell.value <= 1:
                        todayWsData.cell(column=cell.column, row=cell.row).value = 0
                        todayWs.cell(column=cell.column, row=cell.row).value = 0
                except TypeError:
                    continue

        # Check for numbers like 0,001, and if so write "договір є"
        for cells in todayWsData[rangeAM]:
            for cell in cells:
                try:
                    if cell.value > 0 and cell.value <= 1:
                        todayWsData.cell(column=cell.column, row=cell.row).value = str("договір є")
                        todayWs.cell(column=cell.column, row=cell.row).value = str("договір є")
                except TypeError:
                    continue


        # Set 'план є' to those rows wich have 0`s in both columns with conditions
        # Check the range
        list1 = todayTkeWithData.getListOfCellsByCriteria(0, "AS")  
        list2 = todayTkeWithData.getListOfCellsByCriteria(0, "AT")
        columnPlan = columnIndexFromString("AU")

        for cell1 in list1:
            if cell1.value == 0:
                for cell2 in list2:
                    if cell2.value == 0 and cell1.row == cell2.row:
                        cellValueCheck = todayWsData.cell(column=columnPlan, row=cell1.row).value
                        # Check fo cells not to be empty
                        if cellValueCheck == 0 or cellValueCheck == None or cellValueCheck == "":
                            continue
                        else:
                            todayWs.cell(column=columnPlan, row=cell1.row).value = str("план є")
                            todayWs.cell(column=columnPlan+1, row=cell1.row).value = ""
                            todayWs.cell(column=columnPlan+2, row=cell1.row).value = ""
                            todayWs.cell(column=columnPlan+3, row=cell1.row).value = ""
                            todayWs.cell(column=columnPlan+4, row=cell1.row).value = ""
                            todayWs.cell(column=columnPlan+5, row=cell1.row).value = ""
                            todayWs.cell(column=columnPlan+6, row=cell1.row).value = ""

        # Set "інший постачальник" to some companies
        # Opens file with company names and reads them
        with open(self.todayTKE.pathToFile+"\\"+"Інший постачальник.txt", "r") as f:
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
                "Файл ""Інші постачальники"" пустой"\
                + bcolors.ENDC)
            listOtherProvider = []
        else:
            listOtherProvider = content

        rangeIter = "R12" + ":" + "R" + str(numberOfRows)
        columnPlan = columnIndexFromString("AU")
        for cells in todayWsData[rangeIter]:
            for cell in cells:
                if cell.value in listOtherProvider:
                    todayWs.cell(column=columnPlan, row=cell.row).value = "інший постачальник"
                    #todayWsData.cell(column=columnPlan, row=cell.row).value = "інший постачальник"
                    todayWs.cell(column=columnPlan+1, row=cell.row).value = ""
                    todayWs.cell(column=columnPlan+2, row=cell.row).value = ""
                    todayWs.cell(column=columnPlan+3, row=cell.row).value = ""
                    todayWs.cell(column=columnPlan+4, row=cell.row).value = ""
                    todayWs.cell(column=columnPlan+5, row=cell.row).value = ""
                    todayWs.cell(column=columnPlan+6, row=cell.row).value = ""

        # Find the difference between columns with 'план на декаду' and 'поточний лимит'
        for row in range(10, numberOfRows):
            column = columnIndexFromString("AU")
            cellValueCheck = todayWsData.cell(column=column, row=row).value
            if cellValueCheck == None:
                continue
            else:
                columnAS = columnIndexFromString("AS")
                columnAT = columnIndexFromString("AT")
                cellValueCheck1 = todayWsData.cell(column=columnAS, row=row).value
                cellValueCheck2 = todayWsData.cell(column=columnAT, row=row).value
                # Check fo cells not to be empty
                if cellValueCheck1 == 0 and cellValueCheck2 == 0:
                    continue
                elif cellValueCheck1 == None and cellValueCheck2 == None:
                    continue
                else:
                    column1 = columnIndexFromString("BO")
                    value1 = todayWsData.cell(column=column1, row=row).value
                    value2 = todayWsData.cell(column=column, row=row).value
                    try:
                        dx = value1 - value2
                    except TypeError:
                        continue
                    if dx > 1e-6 or dx < -1e-6:
                        column = columnIndexFromString("BW")
                        todayWs.cell(column=column, row=row).value = dx

        self.kyivEnergoMoney(todayTkeWithData)
        #self.oeratorGtsMoney(todayTkeWithData)
        self.calculationOfDebtLikeGarantMM(todayTkeWithData, "Гарант Енерго М ПП")
        self.calculationOfDebtLikeGarantMM(todayTkeWithData, "Нафтогаз Тепло ТОВ")
        todayTkeWithData.close()
        return

    def copyColumn(self):
        
        # Save to tmp variable today and yesterday TKE files
        tmpTodayTKE = self.todayTKE
        tmpYesterdayTKE = self.yesterdayTKE
        # Opens today and yesterday TKE files with pyWin
        self.todayTKE = self.mng.addFileByPath(self.todayTKE.pathToFile, 
                                                self.todayTKE.fileName, 
                                                returnFile=True, 
                                                defaultParser=False, 
                                                openBy=1)
        self.yesterdayTKE = self.mng.addFileByPath(self.yesterdayTKE.pathToFile, 
                                            self.yesterdayTKE.fileName, 
                                            returnFile=True, 
                                            defaultParser=False, 
                                            openBy=1)
        self.todayTKE.open()
        self.yesterdayTKE.open()
        # Set first sheet as active
        todayWs = self.todayTKE.getWs("Sheet1")
        yestWs = self.yesterdayTKE.getWs("Sheet1")
        # Incerts column left to "AS" column in today TKE 
        self.todayTKE.insertColumn("AS")
        # Looks through all rows in today TKE and compare values in "P"
        # column (wich corresponds to company EDRPOU) and if values don`t match
        # then it say`s that there is a new company in today TKE and it should 
        # be copied to yesterday TKE
        column = columnIndexFromString("P")
        numberOfCycles = 0
        startRow = 10
        while True:
            wasMismatch = False
            for row in range(startRow, todayWs.UsedRange.Rows.Count):
                value1 = todayWs.Cells(row, column).Value
                value2 = yestWs.Cells(row, column).Value
                if value1 != value2:
                    wasMismatch = True
                    print(bcolors.OKGREEN +\
                            "Внимание!!! Новое предприятие в списке:", 
                            value1, todayWs.Cells(row, column+2).Value\
                            + bcolors.ENDC)
                    self.yesterdayTKE.insertRow(str(row))
                    for column1 in range(1, yestWs.UsedRange.Columns.Count):
                        # Copy row from today and paste it as values (without formulas) in
                        # yesterday excel book
                        todayWs.Range(str(openpyxl.utils.get_column_letter(column1)) + str(row)).Copy()
                        # Got this from https://docs.microsoft.com/en-us/office/vba/api/excel.xlpastetype
                        # paste values
                        xlPasteValues = -4163
                        yestWs.Range(str(openpyxl.utils.get_column_letter(column1)) + \
                                    str(row)).PasteSpecial(Paste=xlPasteValues)
                    # Copy 1 cell with condition from AT and paste it to AS
                    yestWs.Range("AT" + str(row)).Copy()
                    xlPasteValues = -4163
                    yestWs.Range("AS" + str(row)).PasteSpecial(Paste=xlPasteValues)
                    startRow = row
                    break

            if wasMismatch == False:
                break
            elif wasMismatch == True:
                numberOfCycles += 1
            if numberOfCycles > 5:
                print(bcolors.WARNING +\
                    "Внимание!!! Слишком много несовпадений предприятий со вчерашним днем"\
                    + bcolors.ENDC)
                print(bcolors.WARNING + "Возможна ошибка" + bcolors.ENDC)
        
        # Incerts column "AS" from yesterday TKE into today TKE
        yestWs.Range("AS1:AS2").EntireColumn.Copy()
        todayWs.Paste(todayWs.Range("AS1:AS2"))
        # Saves files with rewriting exsited files in directory
        self.todayTKE.save(self.todayTKE.pathToFile, self.todayTKE.fileNameWithoutExtension, conflictResolution = True)
        self.yesterdayTKE.save(self.yesterdayTKE.pathToFile, self.yesterdayTKE.fileNameWithoutExtension, conflictResolution = True)
        # This somehow closes yesterday TKE also
        self.todayTKE.close()
        self.mng.removeUnCalledFiles()
        # Returns tmp files to variables
        self.todayTKE = tmpTodayTKE
        self.yesterdayTKE = tmpYesterdayTKE

        return

    def restructurization(self, ws, column: int, row: int):
        """Looks through 1730 file and finds company`s debt
        Returns summary debt if its >0, and None if <0

        Keyword arguments:
        ws -- today TKE worksheet
        column -- contract company number
        row -- row company number
        """
        # Set EGRPOU value wich is placed in column right next to 
        # column value
        EGRPOU = ws.cell(column=column+1, row=row).value
        try:
            wsRestr = self.restructurization1730.getWs("Sheet1")
            # Get cell with EGRPOU value in 1730 file
            row = self.restructurization1730.getFirstCellByCriteria(EGRPOU, "D").row
        except AttributeError:
            print(bcolors.OKGREEN +\
                    "В списках договоров реструктуризации " +\
                    "1730 не найдено предприятие с кодом ЕГРПОУ"\
                    + bcolors.ENDC, bcolors.OKGREEN + EGRPOU + bcolors.ENDC)
            return None
        overpaymentColumn = columnIndexFromString("V")
        debtColumn = columnIndexFromString("W")
        overpayment = wsRestr.cell(column=overpaymentColumn, row=row).value
        debt = wsRestr.cell(column=debtColumn, row=row).value
        # Summary of debt (wich are positive values) and overpayment (negative value)
        summary = overpayment + debt
        if summary > 0:
            return summary
        if summary <= 0:
            return None

    def checkRestructurizationData(self, ws, wsData):
        """
        """
        restructWs = self.restructurization1730.getWs()
        rangeIter = "F13" + ":" + "F" + str(restructWs.max_row)
        
        dictOfRestructNumbers = {}
        # Fill dict with EDRPOU and restructurization number from 1730
        for cells in restructWs[rangeIter]:
            for cell in cells:
                if cell.value == None:
                    EDRPOU = str(restructWs.cell(
                                                column=cell.column-2,
                                                row=cell.row).value)
                    number = str(restructWs.cell(
                                                column=cell.column,
                                                row=cell.row+1).value)
                    dictOfRestructNumbers[EDRPOU] = number
                
        rangeIterInner = "P12" + ":" + "P" + str(wsData.max_row)
        for cells in wsData[rangeIterInner]:
            for cell in cells:
                # If EDRPOU in TKE are also in dict from 1730
                # and in TKE there are no data about restructurization
                res = wsData.cell(
                                    column=cell.column-1, 
                                    row=cell.row).value
                if (cell.value != None and
                    str(cell.value) in dictOfRestructNumbers and
                    (res == None or res == "")):
                    # If so then copy restruct number from
                    # 1730 to TKE
                    restructNumber = dictOfRestructNumbers[str(cell.value)]                         
                    ws.cell(column=cell.column - 1,
                            row=cell.row).value = restructNumber
                    wsData.cell(column=cell.column - 1,
                            row=cell.row).value = restructNumber
        
    def kyivEnergoMoney(self, dataFile):
        """Finds kyiv teplo energo money in their passport file
        Keyword arguments:
        dataFile -- TKE file
        """
        try:
            self.kyivEnergoPas.open()
            kyivWs = self.kyivEnergoPas.getWs("Sheet1")
            
            headerColumn = columnIndexFromString("A")
            contractColumn = columnIndexFromString("B")
            moneyColumn = columnIndexFromString("I")
            # Get list with cells with values "Період" in first column ("A")
            listOfHeaders = self.kyivEnergoPas.getListOfCellsByCriteria("Період", "A")
            # Take row of second intro of "Період"
            row = listOfHeaders[1].row
            # From that row iterate untill cell value would contain "рік" and
            # if column next to column "A" not contain "РЗ" in its contract then
            # value in column "I" would be summed up
            money = 0
            while True:
                row += 1
                header = kyivWs.cell(column=headerColumn, row=row).value
                if header == "" or header == None:
                    break
                if "рік" in header:
                    contract = kyivWs.cell(column=contractColumn, row=row).value
                    if "РЗ" not in contract:
                        money += kyivWs.cell(column=moneyColumn, row=row).value
            
            self.kyivEnergoPas.close()
        except (AttributeError, IndexError):
            money = 0
            print(bcolors.WARNING +\
                "Проблема с подсчетом оплаты Київтеплоенерго КП ВО"\
                + bcolors.ENDC)
        
        try:
            # Finds cell with "Київтеплоенерго КП ВО Київради (КМДА)" in TKE file
            kyivEnergoRow = dataFile.getFirstCellByCriteria("Київтеплоенерго " + \
                            "КП ВО Київради (КМДА)", "R").row
            ws = self.todayTKE.getWs("Sheet1")
            wsData = dataFile.getWs()

            # Payment column value
            column=columnIndexFromString("AQ")
            ws.cell(column=column, row=kyivEnergoRow).value = money
            # For all contracts column value
            column=columnIndexFromString("AE")
            column1=columnIndexFromString("AF")
            column2=columnIndexFromString("AD")
            # Set the right value in cell with debt

            realisation = ws.cell(column=column1, row=kyivEnergoRow).value
            remainings = ws.cell(column=column2, row=kyivEnergoRow).value
            
            # Check for not None type
            if realisation == None:
                realisation = 0
            if remainings == None:
                remainings = 0
            
            ws.cell(column=column, row=kyivEnergoRow).value =\
                                    realisation - remainings - money
            wsData.cell(column=column, row=kyivEnergoRow).value =\
                                    realisation - remainings - money
            
            columnAH = columnIndexFromString("AH")
            columnAM = columnIndexFromString("AM")
            # Reopen with pyWin for recalculate formula with debt
            self.todayTKE.save(self.todayTKE.pathToFile, 
                                "tmp", 
                                extension=".xlsx")
            tmpTKE = self.mng.addFileByPath(self.todayTKE.pathToFile, 
                                                "tmp.xlsx", returnFile=True, 
                                                defaultParser=False, 
                                                openBy=1)
            tmpTKE.open()
            tmpWs = tmpTKE.getWs("Sheet1")
            value1 = tmpWs.Cells(kyivEnergoRow, columnAH).Value
            value2 = tmpWs.Cells(kyivEnergoRow, columnAM).Value
            tmpTKE.save(self.todayTKE.pathToFile, 
                        "tmp", 
                        conflictResolution=True)
            tmpTKE.close()
            self.mng.deleteFile(tmpTKE)
            # Make right conditions
            # If have >90% (no debt in column AH) and payed for 
            # restructurization then set 1
            # if other - set 0
            columnAT = columnIndexFromString("AT")

            if (value1 == 0 and
                value2 != None and
                value2 == "договір є"):
                
                wsData.cell(column=columnAT, 
                            row=kyivEnergoRow).value = 1
                ws.cell(column=columnAT, 
                            row=kyivEnergoRow).value = 1
            else:
                wsData.cell(column=columnAT, 
                            row=kyivEnergoRow).value = 0
                ws.cell(column=columnAT, 
                            row=kyivEnergoRow).value = 0


        except AttributeError:
            print(bcolors.WARNING +\
                "Программа не смогла внести данные о задолженности Київтеплоенерго КП ВО"\
                + bcolors.ENDC)
        
        return 

    def oeratorGtsMoney(self, dataFile):
        """Finds operator GTS money in passport file
        Keyword arguments:
        dataFile -- TKE file
        """
        try:
            self.operatorGtsPas.open()
            gtsWs = self.operatorGtsPas.getWs()

            headerColumn = columnIndexFromString("A")
            amountColumn = columnIndexFromString("G")
            costColumn = columnIndexFromString("H")
            paymentColumn = columnIndexFromString("I")
            deptColumn = columnIndexFromString("J")

            # Get list with cells with values "Період" in first column ("A")
            listOfHeaders = self.operatorGtsPas.getListOfCellsByCriteria("Період", "A")
            # Take row of second intro of "Період"
            row = listOfHeaders[1].row
            months = [
                        "липень2019", 
                        "серпень2019", 
                        "вересень2019",
                        "жовтень2019",
                        "листопад2019",
                        "грудень2019"
                    ]
            amount = 0
            cost = 0
            payment = 0
            dept = 0
            while True:
                row += 1
                header = gtsWs.cell(column=headerColumn, row=row).value
                if header == "" or header == None:
                    break
                # Remove spaces because there are from 1 to 4 spaces
                # between month and year (липень 2019 or липень   2019)
                header = str(header).replace(" ", "")
                if "рік" not in header:
                    # For every date in list
                    #for month in months:
                    # If date in passport not equal to this date
                    if header not in months:
                        # Add all values from every column
                        amount += gtsWs.cell(column=amountColumn, 
                                            row=row).value
                        cost += gtsWs.cell(column=costColumn, 
                                            row=row).value
                        payment += gtsWs.cell(column=paymentColumn, 
                                            row=row).value
                        dept += gtsWs.cell(column=deptColumn, 
                                            row=row).value
                    # But if date in passport is equal to that date
                    elif header in months:
                        # Then get payment from that row
                        paymentData = gtsWs.cell(
                                            column=paymentColumn, 
                                            row=row).value
                        # And if that payment is not 0
                        if (paymentData != 0):
                            # Then also add this value
                            payment += paymentData

            self.operatorGtsPas.close()

        except:
            print(bcolors.WARNING +\
                'Программа не смогла посчитать' +\
                    ' деньги Филии "Оператора ГТС" в паспорте' \
                + bcolors.ENDC)

        try:
            # Finds cell with "Філія "Оператор ГТС України"" 
            # using EDRPOU code
            gtsRow = dataFile.getFirstCellByCriteria("42795490", "P").row
            ws = self.todayTKE.getWs()
            wsData = dataFile.getWs()

            # Payment column value
            columnAQ=columnIndexFromString("AQ")
            ws.cell(column=columnAQ, row=gtsRow).value = payment
            # For all contracts column value
            columnAE=columnIndexFromString("AE")
            columnAF=columnIndexFromString("AF")
            columnAD=columnIndexFromString("AD")
            # Set the right value in cell with debt
            realisation = ws.cell(column=columnAF, row=gtsRow).value
            remainings = ws.cell(column=columnAD, row=gtsRow).value
            
            # Check for not None type
            if realisation == None:
                realisation = 0
            if remainings == None:
                remainings = 0

            ws.cell(column=columnAD, row=gtsRow).value = 0
            wsData.cell(column=columnAD, row=gtsRow).value = 0
            ws.cell(column=columnAE, row=gtsRow).value =\
                                                        cost - payment
            wsData.cell(column=columnAE, row=gtsRow).value =\
                                                        cost - payment
            ws.cell(column=columnAF, row=gtsRow).value = cost
            wsData.cell(column=columnAF, row=gtsRow).value = cost
        except:
            print(bcolors.WARNING +\
                'Программа не смогла занести' +\
                    ' деньги Филии "Оператора ГТС"' \
                + bcolors.ENDC)

    def calculationOfDebtLikeGarantMM(self, dataFile, companyName: str):
        """Set right calculation of debt for company 
        like for garant energo M
        Keyword arguments:
        dataFile -- TKE file
        companyName -- name of company (example: "Гарант Енерго М ПП")
        """
        try:
            garantRow = dataFile.getFirstCellByCriteria(companyName, "R").row
            ws = self.todayTKE.getWs("Sheet1")
            # For all contracts column value
            columnAE=columnIndexFromString("AE")
            columnAF=columnIndexFromString("AF")
            # Payment column value
            columnAQ=columnIndexFromString("AQ")
            ws.cell(column=columnAE, row=garantRow).value = \
                            ws.cell(column=columnAF, row=garantRow).value - \
                            ws.cell(column=columnAQ, row=garantRow).value
        except AttributeError:
            print(bcolors.WARNING + \
                "Программа не смогла внести данные о задолженности " + \
                companyName + \
                bcolors.ENDC)

    def hideColumns(self):
        
        listOfNotHiddenColumns = ["B", "C", "D", "P", "Q",
                                    "R", "S", "AD", "AG", "AH",
                                    "AI", "AM", "AS", "AT", "AU",
                                    "BO", "BV", "BW"]
        
        for column in range(1, self.todayTKE.getWs().max_column+1):
            columnLetter = openpyxl.utils.get_column_letter(column)
            if columnLetter not in listOfNotHiddenColumns:
                self.todayTKE.getWs().column_dimensions[columnLetter].hidden = True
        return

    def addFilter(self):
        FullRange = "A9:" + openpyxl.utils.get_column_letter(
                                self.todayTKE.getWs().max_column) + \
                                str(self.todayTKE.getWs().max_row)
        self.todayTKE.getWs().auto_filter.ref = FullRange


    def generateName(self):
        """Generates name for file TKE_ПСО
        with current date and month
        Returns string without file extension
        """
        day = datetime.datetime.today().day
        month = datetime.datetime.today().month
        year = datetime.datetime.today().year
        monthInRussian = [r"январь", r"февраль", r"март", 
                            r"апрель", r"май", r"июнь", 
                            r"июль", r"август", r"сентябрь", 
                            r"октябрь", r"ноябрь", r"декабрь"]
        fileName = "90%ТКЕ_ПСО_" + monthInRussian[month-1]
        fileName += "(" + str(day) + "."
        if month < 10:
            fileName += "0" + str(month) + "."
        elif month >= 10:
            fileName += str(month) + "."
        fileName += str(year) + ")"
        return fileName


class TKELess(TKE):

    def prepareTodayTkeFile(self):
        
        tmpTke = self.todayTKE
        self.todayTKE = self.mng.addFileByPath(self.todayTKE.pathToFile, 
                                            self.todayTKE.fileName, 
                                            returnFile=True, 
                                            defaultParser=False, 
                                            openBy=1)
        self.todayTKE.open()
        ws = self.todayTKE.getWs("Sheet1")

        # Delete one row with numbers of column
        ws.Rows(10).Delete()

        # Translate column names with different contracts in numbers
        columns = ["E", "G", "I", "K", "M", "N"]
        columnList = []
        for col in columns:
            columnList.append(columnIndexFromString(col))
        
        numberOfDeletedCompanies = 0
        for row in range(11, ws.UsedRange.Rows.Count):
            # Move 1 row up because excel moves entire sheet up
            # by number of deleted rows after deleting row
            row -= numberOfDeletedCompanies
            # Get data from columns
            # the data is the contracts name
            haveContract = True
            # Check if this is a row with company
            # and not the row with region
            if ws.Cells(row, 3).Value != None:
                for col in columnList:
                    value = ws.Cells(row, col).Value
                    
                    # Check if there is at least one contract
                    if value == None or value == "":
                        haveContract = False
                    else:
                        haveContract = True
                        break
            
            # If that company haven`t contract, then delete row with
            # that company
            if haveContract == False:
                ws.Rows(row).Delete()
                numberOfDeletedCompanies += 1

        self.todayTKE.save(self.todayTKE.pathToFile, 
                        self.todayTKE.fileNameWithoutExtension, 
                        conflictResolution = True)
        self.todayTKE.close()
        self.mng.removeFile(self.todayTKE)
        self.todayTKE = tmpTke
        

    def copyColumn(self):
        """This is just copy of copyColumn from TKE class
        with some changes highlited with "" coments
        """

        # Save to tmp variable today and yesterday TKE files
        tmpTodayTKE = self.todayTKE
        tmpYesterdayTKE = self.yesterdayTKE
        # Opens today and yesterday TKE files with pyWin
        self.todayTKE = self.mng.addFileByPath(self.todayTKE.pathToFile, 
                                                self.todayTKE.fileName, 
                                                returnFile=True, 
                                                defaultParser=False, 
                                                openBy=1)
        self.yesterdayTKE = self.mng.addFileByPath(self.yesterdayTKE.pathToFile, 
                                                    self.yesterdayTKE.fileName, 
                                                    returnFile=True, 
                                                    defaultParser=False, 
                                                    openBy=1)
        self.todayTKE.open()
        self.yesterdayTKE.open()
        # Set first sheet as active
        todayWs = self.todayTKE.getWs("Sheet1")

        """Here is the change of sheet name"""
        yestWs = self.yesterdayTKE.getWs("ТКЕ1920")

        # Incerts column left to "AS" column in today TKE 
        self.todayTKE.insertColumn("AS")
        # Delete column "B" from yesterday
        yestWs.Columns("B").Delete()

        yesterdayLenght = yestWs.UsedRange.Rows.Count
        todayLenght = todayWs.UsedRange.Rows.Count
        limit = todayLenght if todayLenght > yesterdayLenght else yesterdayLenght
        
        # Looks through all rows in today TKE and compare values in "P"
        # column (wich corresponds to company EDRPOU) and if values don`t match
        # then it say`s that there is a new company in today TKE and it should 
        # be copied to yesterday TKE
        """Added defferent column name in yesterday TKE"""
        columnEDRPOU = columnIndexFromString("P")
        columnCompany = columnIndexFromString("R")
        
        for row in range(10, limit):            
            if (todayWs.Cells(row, columnEDRPOU).Value != 
                    yestWs.Cells(row, columnEDRPOU).Value):
                if (todayWs.Cells(row+1, columnEDRPOU).Value == 
                        yestWs.Cells(row, columnEDRPOU).Value):
                    # Company added
                    print(bcolors.OKGREEN +\
                            "Новое предприятие в списке:", 
                            todayWs.Cells(row, columnCompany).Value\
                            + bcolors.ENDC)
                    self.yesterdayTKE.insertRow(str(row))
                    for column in range(1, yestWs.UsedRange.Columns.Count):
                        # Copy row from today and paste it as 
                        # values (without formulas) in yesterday 
                        # excel book
                        todayWs.Range(openpyxl.utils.get_column_letter(column)\
                            + str(row)).Copy()
                        # Got this from 
                        # https://docs.microsoft.com/en-us/office/vba/api/excel.xlpastetype
                        # paste values
                        xlPasteValues = -4163
                        yestWs.Range(openpyxl.utils.get_column_letter(column) + \
                                    str(row)).PasteSpecial(Paste=xlPasteValues)
                    # Copy 1 cell with condition from AT 
                    # and paste it to AS
                    """Renamed columns in yesterdayTKE"""
                    yestWs.Range("AT" + str(row)).Copy()
                    xlPasteValues = -4163
                    yestWs.Range("AS" + str(row)).PasteSpecial(Paste=xlPasteValues)

                elif (todayWs.Cells(row, columnEDRPOU).Value == 
                        yestWs.Cells(row+1, columnEDRPOU).Value):
                    # Company deleted
                    print(bcolors.OKGREEN +\
                            "Предприятие исчезло из списка:", 
                            yestWs.Cells(row, columnCompany).Value\
                            + bcolors.ENDC)
                    yestWs.Rows(row).Delete()
        
        # Incerts column "AS" from yesterday TKE into today TKE
        """Renamed columns in yesterdayTKE"""
        yestWs.Range("AS1:AS2").EntireColumn.Copy()
        todayWs.Paste(todayWs.Range("AS1:AS2"))
        # Saves files with rewriting exsited files in directory
        self.todayTKE.save(self.todayTKE.pathToFile, 
                            self.todayTKE.fileNameWithoutExtension, 
                            conflictResolution = True)
        self.yesterdayTKE.save(self.yesterdayTKE.pathToFile, 
                            self.yesterdayTKE.fileNameWithoutExtension, 
                            conflictResolution = True)
        # This somehow closes yesterday TKE also
        self.todayTKE.close()
        self.mng.removeUnCalledFiles()
        # Returns tmp files to variables
        self.todayTKE = tmpTodayTKE
        self.yesterdayTKE = tmpYesterdayTKE
    
    def run(self):
        self.prepareTodayTkeFile()
        self.copyColumn()
        self.mainCalculations()
        name = self.generateName()
        self.hideColumns()
        self.todayTKE.save(self.todayTKE.pathToFile, 
                            name, extension=".xls")
        self.deleteFiles()
