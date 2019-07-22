import openpyxl


class ExcelBook:

    def __init__(self, name: str, data_only=False):
        self.wb = openpyxl.load_workbook(name, data_only=data_only)
        self.ws = self.wb[self.wb.sheetnames[0]]

    def save(self, name: str):
        self.wb.save(name)
        self.wb.close()
        return
    
    def close(self):
        self.wb.close()
        return 

    def setCellsInColumnByRowCoord(self, rowCoord: int, sheetRange: str, value):
        # Finds values in one column (sheetRange) by row coordinate (rowCoord)
        # and then set value of that cell to (value)
        for cells in self.ws[sheetRange]:
            for cell in cells:
                if cell.row == rowCoord:
                    cell.value = value
        return

    def setCellsInRowByColumnCoord(self, columnCoord: int, sheetRange: str, value):
        # Finds values in one row (sheetRange) by column coordinate (columnCoord)
        # and then set value of that cell to (value)
        for cells in self.ws[sheetRange]:
            for cell in cells:
                if cell.column == rowCoord:
                    cell.value = value
        return

    def findCellByStr(self, str: str):
        # Finds first cell by searchin in whole sheet the target (str)
        for cells in self.ws:
            for cell in cells:
                if cell.value == str:
                    return cell
        return Null

    def fillCellsWithString(self, cell, str: str):
        cell.value = str
        return

    def getListOfCellsWithCriteria(self, range: str, criteria):
        listOfCells = []

        if  range.isdigit() == True:
            for cells in self.ws[range]:
                for cell in cells:
                    if cell.value == criteria:
                        listOfCells.append(cell)
        elif range.isdigit() == False:
            for cell in self.ws[range]:
                if cell.value == criteria:
                    listOfCells.append(cell)
        return listOfCells


class TKE:

    def __init__(self, wbName: str):
        self.wb = ExcelBook(wbName, data_only=True)
        self.header = self.wb.ws['A1':'BU9']

    def start(self):
        """Deleting first row"""
        #self.wb.ws.delete_rows(1)    #Check the row number

        """Set 1 and 'договир э' to those companies who have a restructurization contract"""
        for cells in self.wb.ws['F12:F22']:
            for cell in cells:
                if cell.value != None:
                    self.wb.ws[str('H')+str(cell.row)] = 1
                    self.wb.ws[str('R')+str(cell.row)] = str('договир э')

        """Opens ysterday workbook and copies one specific column to current 
        workbook shifting other columns
        """
        self.wbFromYesterday = ExcelBook('./second.xlsx', data_only=True)
        if self.wb.ws.max_row == self.wbFromYesterday.ws.max_row:       #TODO make the right check
            self.wb.ws.move_range('I1:R22', rows=0, cols=1)           #Check the right range!!!!!!!!!
            for cell in self.wbFromYesterday.ws['I']:
                self.wb.ws.cell(row = cell.row, column = cell.column, value = cell.value)
        else:
            raise Exception('Different number of rows in both docs\n. The first has: {}'.format(self.wb.ws.max_row))
        self.wbFromYesterday.close()

        """Transfer data from 'поточний лимит' to 'попередний лимит'"""
        for cells1, cells2 in zip(self.wb.ws['P12:R22'], self.wb.ws['M12:O22']):
            for cell1, cell2 in zip(cells1, cells2):
                if cell1.row == cell2.row:
                    self.wb.ws.cell(column=cell2.column, row=cell2.row, value=cell1.value)
        
        """Multiply cells in column 'план на декаду' by 3"""
        for cells in self.wb.ws['J12:L22']:
            for cell in cells:
                self.wb.ws.cell(column=cell.column, row=cell.row, value=cell.value*3)
        
        """Set 'план э' to those rows wich have 0`s in both columns with conditions"""
        list1 = self.wb.getListOfCellsWithCriteria('H', 0)             # Check the range
        list2 = self.wb.getListOfCellsWithCriteria('I', 0)             # Check the range
        
        for cell1 in list1:
            if cell1.value == 0:
                for cell2 in list2:
                    if cell2.value == 0 and cell1.row == cell2.row:
                        self.wb.ws[str('J')+str(cell1.row)] = str('план э')
                        self.wb.ws[str('K')+str(cell1.row)] = 0
                        self.wb.ws[str('L')+str(cell1.row)] = 0

        """Find the difference between columns with 'план на декаду' and 'поточний лимит'"""
        for row in range(12, self.wb.ws.max_row):
            if self.wb.ws[str('J')+str(row)].value != 'план э':
                self.wb.ws.cell(column=openpyxl.utils.column_index_from_string(str('T')),
                                row=row,
                                value=self.wb.ws[str('P')+str(row)].value-self.wb.ws[str('J')+str(row)].value
                                )

        self.wb.save('./out.xlsx')
        


class FormulaTransformer:

    def __init__(self, ws):
        self.ws = ws

    def findColumnsWithFormulas(self):
        """Creates a list with coordinates of all cells wich
        contains ranges (=SUM(A1:A3), =A1+B1, but not =RANDOM(1:100))
        from upper cells in the search range (very usefull to find
        under the header of the table because its often the place
        where formulas are created and then broached to the bottom
        of the table)

        Keyword arguments:
        ws -- active worksheet
    
        """
        columnsWithFormulas = []
        for cells in ws['B12':'L12']:
            for cell in cells:
                token = openpyxl.formula.Tokenizer(str(cell.value))
                for element in token.items:
                    if element.subtype == 'RANGE':
                        columnsWithFormulas.append(cell.coordinate)

        """Deleting repeats(multiple inclusions)"""
        columnsWithFormulasWithoutRepeat = []
        for i in columnsWithFormulas:
            if i not in columnsWithFormulasWithoutRepeat:
                columnsWithFormulasWithoutRepeat.append(i)

        return columnsWithFormulasWithoutRepeat

    def checkRangeInFormulas(self, range: str, columnsWithFormulas):
        """Check a list with coordinates wich contains forulas with 
        some kind of ranges for intersection with modified 
        range after insertion of new column
        This function shoul be used after insertion of the new column

        Keyword arguments:
        ws -- active worksheet
        range -- string with modified range (should be 
        up to right corner of table)
        columnsWithFormulas -- list of all top (under the header) 
        cells with formulas wich contains ranges

        """
        columnsList = []
        split = range.split(':')
        minCoordinate = openpyxl.utils.coordinate_to_tuple(split[0])
        maxCoordinate = openpyxl.utils.coordinate_to_tuple(split[1])
        minColumn, maxColumn = minCoordinate[1], maxCoordinate[1]
        for coordinate in columnsWithFormulas:
            column = openpyxl.utils.coordinate_to_tuple(coordinate)[1]
            if column < minColumn:
                columnsList.append(openpyxl.utils.get_column_letter(column))
        return columnsList

    def modifyFormulaInColumnFterInsertion(self, targetColumns, newColumnCoordinate: str):
        """Function looks throug whole column and if there is a cell with formula
        like =A1+B1 and if column has been inserted between A1 and B1 then it will
        modify formula to =A1+C1

        Keyword arguments:
        ws -- active worksheet
        targetColumns -- sequence of columns that should be modified
        newColumnCoordinates -- integer that corresponds to number of inserted columns

        """
        for column in targetColumns:
            print('targetColumns = ', targetColumns)
            for cell in ws[column]:
                print('beginning of cycle', cell.value)
                token = openpyxl.formula.Tokenizer(str(cell.value))
                if cell.value == None:
                    continue
                newCellValue = str('=')
                for element in token.items:
                    print('beginning if subsycle', 'element.value = ', element.value,
                         'element.type = ', element.type, 'element.subtype = ', element.subtype)
                    if element.subtype == 'RANGE':
                        coord = openpyxl.utils.coordinate_to_tuple(element.value)
                        print('coord = ', coord)
                        newColumn = coord[1] + newColumnCoordinate
                        print('newColumn = ', newColumn)
                        newCoordinate = str(openpyxl.utils.get_column_letter(newColumn) + str(coord[0]))
                        newCellValue+=newCoordinate
                    else:
                        newCellValue+=element.value
                    print('newCellValue = ', newCellValue)
                    print('exit subloop')
                cell.value = newCellValue
                print('exit loop')
                print(cell.value)
        return


def updateCell(coord, activeSheet):
    for cells in activeSheet['H12:H22']:
        for cell in cells:
            if cell.row == coord:
                cell.value = 1
    return

def restructurization(activeSheet):
    for cells in activeSheet['F12':'F22']:
        for cell in cells:
            if cell.value != None:
                updateCell(cell.row, activeSheet)
    return
 
def findColumnsWithFormulas(ws):
    """Creates a list with coordinates of all cells wich
    contains ranges (=SUM(A1:A3), =A1+B1, but not =RANDOM(1:100))
    from upper cells in the search range (very usefull to find
    under the header of the table because its often the place
    where formulas are created and then broached to the bottom
    of the table)

    Keyword arguments:
    ws -- active worksheet
    
    """
    columnsWithFormulas = []
    for cells in ws['B12':'L12']:
        for cell in cells:
            token = openpyxl.formula.Tokenizer(str(cell.value))
            for element in token.items:
                if element.subtype == 'RANGE':
                    columnsWithFormulas.append(cell.coordinate)

    """Deleting repeats(multiple inclusions)"""
    columnsWithFormulasWithoutRepeat = []
    for i in columnsWithFormulas:
        if i not in columnsWithFormulasWithoutRepeat:
            columnsWithFormulasWithoutRepeat.append(i)

    return columnsWithFormulasWithoutRepeat

def checkRangeInFormulas(ws, range: str, columnsWithFormulas):
    """Check a list with coordinates wich contains forulas with 
    some kind of ranges for intersection with modified 
    range after insertion of new column
    This function shoul be used after insertion of the new column

    Keyword arguments:
    ws -- active worksheet
    range -- string with modified range (should be 
    up to right corner of table)
    columnsWithFormulas -- list of all top (under the header) 
    cells with formulas wich contains ranges

    """
    columnsList = []
    split = range.split(':')
    minCoordinate = openpyxl.utils.coordinate_to_tuple(split[0])
    maxCoordinate = openpyxl.utils.coordinate_to_tuple(split[1])
    minColumn, maxColumn = minCoordinate[1], maxCoordinate[1]
    for coordinate in columnsWithFormulas:
        column = openpyxl.utils.coordinate_to_tuple(coordinate)[1]
        if column < minColumn:
            columnsList.append(openpyxl.utils.get_column_letter(column))
    return columnsList

def modifyFormulaInColumnFterInsertion(ws, targetColumns, newColumnCoordinate: str):
    """Function looks throug whole column and if there is a cell with formula
    like =A1+B1 and if column has been inserted between A1 and B1 then it will
    modify formula to =A1+C1

    Keyword arguments:
    ws -- active worksheet
    targetColumns -- sequence of columns that should be modified
    newColumnCoordinates -- integer that corresponds to number of inserted columns

    """
    for column in targetColumns:
        print('targetColumns = ', targetColumns)
        for cell in ws[column]:
            print('beginning of cycle', cell.value)
            token = openpyxl.formula.Tokenizer(str(cell.value))
            if cell.value == None:
                continue
            newCellValue = str('=')
            for element in token.items:
                print('beginning if subsycle', 'element.value = ', element.value,
                     'element.type = ', element.type, 'element.subtype = ', element.subtype)
                if element.subtype == 'RANGE':
                    coord = openpyxl.utils.coordinate_to_tuple(element.value)
                    print('coord = ', coord)
                    newColumn = coord[1] + newColumnCoordinate
                    print('newColumn = ', newColumn)
                    newCoordinate = str(openpyxl.utils.get_column_letter(newColumn) + str(coord[0]))
                    newCellValue+=newCoordinate
                else:
                    newCellValue+=element.value
                print('newCellValue = ', newCellValue)
                print('exit subloop')
            cell.value = newCellValue
            print('exit loop')
            print(cell.value)
    return

"""
# Opening workbook at sheet 1
wb = openpyxl.load_workbook('./first.xlsx')
activeSheet = wb[wb.sheetnames[0]]

# Set 1 in column AS
restructurization(activeSheet)
# Check total row number
totalRowNumber = activeSheet.max_row

# Open second workbook at sheet 1
wb2 = openpyxl.load_workbook('./second.xlsx')
activeSheet2 = wb2[wb2.sheetnames[0]]


columns = findColumnsWithFormulas(activeSheet)
activeSheet.move_range('K12:N22', rows=0, cols=1, translate=True)
columnsList = checkRangeInFormulas(activeSheet, 'K12:N22', columns)
modifyFormulaInColumnFterInsertion(activeSheet, columnsList, 1)



wb.save('out.xlsx')
wb.close()
"""

first = TKE('./first.xlsx')
first.start()