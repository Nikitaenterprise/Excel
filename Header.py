import openpyxl
from ExcelBook import ExcelBook

class Header(ExcelBook):

    def __init__(self, name: str, headerRange: str):
        ExcelBook.__init__(self, name, data_only=True)
        self.header = headerRange
        split = headerRange.split(":")
        if len(split) < 2:
            raise Exception("Header can`t be 1 cell {}".format(headerRange))
        self.leftTopCoordinate = headerRange.split(":")[0]
        self.rightBotCoordinate = headerRange.split(":")[1]
        self.leftTopRow = openpyxl.utils.coordinate_to_tuple(self.leftTopCoordinate)[0]
        self.leftTopColumn = openpyxl.utils.coordinate_to_tuple(self.leftTopCoordinate)[1]
        self.rightBotRow = openpyxl.utils.coordinate_to_tuple(self.rightBotCoordinate)[0]
        self.rightBotColumn = openpyxl.utils.coordinate_to_tuple(self.rightBotCoordinate)[1]

    #def generator(self):
    #    """Generator throug the header.
    #    Returns cell
    #    """
    #    __super__(self, )
    #    for cells in self.wb.ws[self.header]:
    #        for cell in cells:
    #            yield cell

    def generator(self):
        """Generator throug the header.
        Returns cell
        """
        return super().generator(header)

    def findCellByStr(self, value: str):
        """Finds first cell by searchin in the header 
        the target value. Returns a list [row, column]

        Keyword argument:
        value -- searching string
        """
        for cells in self.wb.ws[self.header]:
            for cell in cells:
                if cell.value == value:
                    return [cell.row, cell.column]

    def unmergeCells(self):
        
        return