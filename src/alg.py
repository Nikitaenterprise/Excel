import os
import datetime
import time
from copy import copy

import openpyxl
import win32com.client

from src.manager import *

# Just for less writning
columnIndexFromString = openpyxl.utils.column_index_from_string
    
class Algorithm():
    def __init__(self, dir, files):
        self.start_time = time.time()
        self.mng = Manager(os.path.abspath(dir))
        self.numberOfFilesToStart = files
        self.checkIfDirectoryIsReady(dir)
    

    def checkIfDirectoryIsReady(self, path: str):
        pass

    def deleteFiles(self, programmIsDone=True):
        pass

    def run(self):
        pass

    def getTimeOfRun(self):
        return time.time() - self.start_time

class bcolors:
    HEADER = ""#' \033[95m '
    OKBLUE = ""#' \033[94m '
    OKGREEN = ""#' \033[92m '
    WARNING = ""#' \033[93m '
    FAIL = ""#' \033[91m '
    ENDC = ""#' \033[0m '
    BOLD = ""#' \033[1m '
    UNDERLINE = ""#' \033[4m '