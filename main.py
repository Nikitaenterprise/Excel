import os

from src.plan import FiscalPlan
from src.tke import TKE

def killProcess(hardKill: int):
    processes = os.popen('tasklist').readlines()
    for process in processes:
        if "EXCEL.EXE" in process:
            if hardKill == 0:
                os.system("taskkill /im EXCEL.EXE")
            if hardKill == 1:
                os.system("taskkill /f /im EXCEL.EXE")

if __name__ == "__main__":
    killProcess(0)
    print("Программа закрывает приложение excel")
    print("Сохраните книги, если они были открыты\n")
    print("Введите:\n1 для фин-плана\n2 для ТКЕ_ПСО")
    while True:
        what = input()

        if what == "1":
            killProcess(1)
            fp = FiscalPlan(r"FiscalPlan", 6)
            fp.run()
            killProcess(0)
            break
        elif what == "2":
            killProcess(1)
            tke = TKE(r"TKE", 4)
            tke.run()
            killProcess(0)
            break
        else:
            print("Неправильный ввод")
    print("Программа завершила работу")
    input()


