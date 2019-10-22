from src.alg import *

class NKREKU2(Algorithm):

    def checkIfDirectoryIsReady(self):
        self.mng.addFilesInDir()

        self.outFile = self.mng.getFile("Ф2", extension=".xlsx")
        self.outFile.shouldBeDeleted = False

        self.mng.getFile("Оборотно-сальдова вiдомiсть")

        self.mng.deleteUnCalledFiles()               
        self.mng.allFromXlsToXlsx()

        try:
            self.saldo = self.mng.getFile("Оборотно-сальдова вiдомiсть", extension=".xlsx")
            
            if self.mng.getNumberOfFiles() != self.numberOfFilesToStart:
                raise AttributeError
        except AttributeError:
            print("Не хватает файлов для работы. Проверьте директорию " + str(path))
            msg = r"""Файлы, нужные для работы: 
            После исправления запустите программу заново. Сейчас программа завершит работу
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
        self.getValues()
        self.outFile.save(self.outFile.pathToFile, "НКРЕКП №2", extension=".xlsx")
        self.deleteFiles()

    def getValues(self):

        se