import openpyxl


class ExcelBook:

    def __init__(self, name: str):
        self.wb = openpyxl.load_workbook(name)
        self.ws = wb[wb.sheetnames[0]]

    def getMaxRowForSheet(self, ws):
        return ws.max_row

    def getMaxColForSheet(self, ws):
        return ws.max_col

    def save(self, name: str):
        self.wb.save(name)
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


class TKE:

    def __init__(self, wbName: str):
        self.excelBook = ExcelBook(wbName)
        self.header = excelBook.ws['A1':'BU9']

    def restructurization(self):
        for cells in excelBook.ws['H12:H22']:
            for cell in cells:
                if cell.value != None:
                    excelBook.updateCell(cell.row, excelBook.ws)
        return

    def start(self):
        restructurization()
        


def updateCell(coord, activeSheet):
    for cells in activeSheet['H12:H22']:
        for cell in cells:
            if cell.row == coord:
                cell.value = 1
    return

def restructurization(activeSheet):
    for cellObj in activeSheet['F12':'F22']:
        for cell in cellObj:
            if cell.value != None:
                updateCell(cell.row, activeSheet)
    return
 
def updateFormula(activeSheet, inertedColumn):
    for cells in activeSheet['B12':'L12']:
        for cell in cells:
            token = openpyxl.formula.Tokenizer(str(cell.value))
            #print(cell.value, token.items)
            for element in token.items:
                if element == 'OPERAND RANGE':
                    token.value
                print(element)
                

    return


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


updateFormula(activeSheet, 1)
#activeSheet.insert_cols(openpyxl.utils.column_index_from_string('H'), fill_formulae=False)
#for cell in activeSheet2['I']:
#   activeSheet.cell(row = cell.row, column = cell.column-1, value = cell.value)

#activeSheet.move_range('J12:L22', rows=0, cols=1, translate=True)

#for cell in activeSheet['J']:
#    oldColumn = openpyxl.utils.get_column_letter(cell.column)
#    newColumn = openpyxl.utils.get_column_letter(cell.column + 1)
#    newAdress = str(newColumn) + str(cell.row)
    
    #print(oldColumn, str(cell.row), cell.value, newColumn, newAdress, activeSheet[str(oldColumn)+str(cell.row)].value)
    # updateFormula(cell.value)
    #activeSheet[newAdress] = openpyxl.formula.translate.Translator(cell.value, origin=str(oldColumn)+str(cell.row)).translate_formula(newAdress)
    #print(oldColumn, str(cell.row), cell.value, newColumn, newAdress)


wb.save('out.xlsx')
wb.close()