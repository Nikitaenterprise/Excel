import openpyxl


class ExcelBook:

    def __init__(self, name: str, data_only=False):
        self.wb = openpyxl.load_workbook(name, data_only=data_only)
        self.ws = self.wb[self.wb.sheetnames[0]]

    def save(self, name: str):
        """Closes file and save it to project root dir
        with a given name

        Keyword arguments:
        name -- name of saved file

        """
        self.wb.save(name)
        self.wb.close()
        return
    
    def close(self):
        """Closes file without saving"""
        self.wb.close()
        return 

    def setCellsInColumnByRowCoord(self, row: int, column: str, value):
        """Finds values in one column by row coordinate
        and then set value of that cell

        Keyword arguments:
        row -- row number
        column -- column in which the searh would happend
        value -- this would be set to the cell

        """
        for cells in self.ws[column]:
            for cell in cells:
                if cell.row == row:
                    cell.value = value
        return

    def setCellsInRowByColumnCoord(self, row: int, column: str, value):
        """Finds values in one row by column coordinate
        and then set value of that cell

        Keyword arguments:
        row -- row number
        column -- column in which the searh would happend
        value -- this would be set to the cell

        """
        for cells in self.ws[str(row)]:
            for cell in cells:
                if cell.column == openpyxl.utils.column_index_from_string(column):
                    cell.value = value
        return

    def findCellByStr(self, value):
        """Finds first cell by searchin in whole sheet 
        the target value

        Keyword argument:
        value -- searching value (str, int, ...)

        """
        for cells in self.ws:
            for cell in cells:
                if cell.value == value:
                    return cell
        return Null

    def getListOfCellsWithCriteria(self, range: str, criteria):
        """Returns list of cells with values equal to criteria

        Keyword argument:
        range -- search range, could be with or 
        without ':' (I22:J22, or I)
        criteria -- search criteria in cells values

        """
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

first = TKE('./first.xlsx')
first.start()