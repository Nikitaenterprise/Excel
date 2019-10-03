from src.plan import FiscalPlan
from src.tke import TKE

import openpyxl
import win32com.client

if __name__ == "__main__":
    print("Введите:\n1 для фин-плана\n2 для ТКЕ_ПСО")
    while True:
        what = input()
        if what == "1":
            TKE = FiscalPlan(r"FiscalPlan")
            TKE.run()
            break
        elif what == "2":
            tke = TKE(r"TKE")
            tke.run()
            break
        else:
            print("Не правильный ввод")
    print("Программа завершила работу")
    input()