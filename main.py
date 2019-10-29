import os
import traceback

from src.plan import FiscalPlan
from src.tke import TKE
from src.decade import Decade
from src.nkreku2 import NKREKU2
from src.nkreku_pat import NKREKU_PAT

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
    print("Введите:")
    print("\t1 для фин-плана")
    print("\t2 для ТКЕ_ПСО")
    print("\t3 для форм 1-8")
    print("\t4 для формы НКРЭКУ №2")
    print("\t5 для формы НКРЭКУ ВТВ+НОРМ")

    while True:
        what = input()
        killProcess(1)
        try:
            if what == "1":
                fp = FiscalPlan(r"FiscalPlan", 6)
                fp.run()
                print("Время выполнения :", fp.getTimeOfRun())
                killProcess(0)
                break
            elif what == "2":
                tke = TKE(r"TKE", 4)
                tke.run()
                print("Время выполнения :", tke.getTimeOfRun())
                killProcess(0)
                break
            elif what == "3":
                decade = Decade(r"Decade", 12)
                decade.run()
                print("Время выполнения :", decade.getTimeOfRun())
                killProcess(0)
                break
            elif what == "4":
                nkreku2 = NKREKU2(r"NKREKU2", 2)
                nkreku2.run()
                print("Время выполнения :", nkreku2.getTimeOfRun())
                killProcess(0)
                break
            elif what == "5":
                nkreku_pat = NKREKU_PAT(r"NKREKU_PAT", 2)
                nkreku_pat.run()
                print("Время выполнения :", nkreku_pat.getTimeOfRun())
                killProcess(0)
                break
            else:
                print("Неправильный ввод")
        except Exception:
           print("Возникло необработанное исключение")
           print(traceback.format_exc())
    print("Программа завершила работу")
    input()