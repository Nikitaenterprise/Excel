import os


class Directory():

    def __init__(self, pathToDir: str, listOfPatterns: list, numberOfFiles: int):
        self.path = pathToDir
        self.patterns = listOfPatterns
        self.numberOfFiles = numberOfFiles
        return

    def checkIfDirectoryIsReady(self):
        numberOfFiles = self.scanDirectory()
        # Check the dir for needed files
        while True:
            if numberOfFiles == self.numberOfFiles:
                break
            if numberOfFiles > 6:
                print("Слишком много экселевских файлов в папке")
                print("Должно быть ровно"+numberOfFiles)
                print("Программа пробует удалить ненужные")
                self.deleteFiles()
            numberOfFiles = self.scanDirectory()
        return

    def scanDirectory(self, ):
        """Scans the directory with os.walk() for excel files
        and set class excel book variables for folowing work
        """
        #print(os.path.abspath(self.path))
        numberOfFiles = 0
        # r=root, d=directories, f = files
        for r, d, f in os.walk(self.path):
            for file in f:
                if ".xls" in file or ".xlsx" in file:
                    numberOfFiles += 1

        return numberOfFiles

    def deleteFiles(self):

        return

findPatterns = ["Прогнозне надходження", "ЗБУТ", "ПАТ", "НаКР", "ТЕЦ"]
reader = Directory(r"src", findPatterns, 6)