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