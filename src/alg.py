import os
import datetime
from copy import copy

import openpyxl
import win32com.client

from src.manager import *

class Algorithm():
    def __init__(self, dir, files):
        self.mng = Manager(os.path.abspath(dir))
        self.numberOfFilesToStart = files
        self.checkIfDirectoryIsReady(dir)

    def checkIfDirectoryIsReady(self, path: str):
        pass

    def deleteFiles(self, programmIsDone=True):
        pass

    def run(self):
        pass

class bcolors:
    HEADER = ' \033[95m '
    OKBLUE = ' \033[94m '
    OKGREEN = ' \033[92m '
    WARNING = ' \033[93m '
    FAIL = ' \033[91m '
    ENDC = ' \033[0m '
    BOLD = ' \033[1m '
    UNDERLINE = ' \033[4m '