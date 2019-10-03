import openpyxl
import win32com.client

from src.manager import *


class Decade:
    def __init__(self, dir: str):
        self.mng = Manager(os.path.abspath(dir))
        self.numberOfFilesToStart = 0