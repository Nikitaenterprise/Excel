import openpyxl
import re

class ExcelScripts:

    def __init__(self, name: str):
        self.wb = openpyxl.load_workbook(name)
        self.ws = wb[wb.sheetnames[0]]

    def save(self, name: str):
        self.wb.save(name)
        self.wb.close()
        return

    def setCellsInColumnByRowCoord(self, rowCoord: int, sheetRange: str, value):
        # Finds 
        for cells in self.ws[sheetRange]:
            for cell in cells:
                if cell.row == rowCoord:
                    cell.value = value
        return

    def setCellsInRowByColumnCoord(self, columnCoord: int, sheetRange: str, value):
        for cells in self.ws[sheetRange]:
            for cell in cells:
                if cell.column == rowCoord:
                    cell.value = value
        return


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
                print(cell, cell.coordinate, cell.value)
                updateCell(cell.row, activeSheet)
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

activeSheet.insert_cols(openpyxl.utils.column_index_from_string('I'))
for cell in activeSheet2['I']:
   activeSheet.cell(row = cell.row, column = cell.column, value = cell.value, )


wb.save('out.xlsx')
wb.close()