import os
import traceback
import time

from src.plan import FiscalPlan
from src.tke import TKE, TKELess
#from src.decade import Decade
from src.decadeV2 import Decade
from src.nkreku_pat import NKREKU_PAT
from src.nkreku_pat_zbut import NKREKU_PAT_ZBUT_VTV_naselenie
from src.nkreku2 import NKREKU2
from src.nkreku3 import NKREKU3
from src.nkreku5 import NKREKU5

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
    print("\t20 для ТКЕ_ПСО без неработающих предприятий")
    print("\t3 для форм 1-8")
    print("\t4 для формы НКРЭКУ №2")
    print("\t5 для формы НКРЭКУ №3")
    print("\t6 для формы НКРЭКУ №5")
    print("\t7 для формы НКРЭКУ ПАТ(ВТВ+НОРМ) месячная")
    print("\t8 для форм НКРЭКУ ПАТ_ЗБУТ(население + ВТВ+НОРМ)")

    while True:
        what = input()
        killProcess(1)
        time.sleep(1)
        try:
            if what == "1":
                alg = FiscalPlan(r"FiscalPlan", 6)
                alg.run()
                print("Время выполнения :", alg.getTimeOfRun())
                killProcess(0)
                break
            elif what == "2":
                alg = TKE(r"TKE", 4)
                alg.run()
                print("Время выполнения :", alg.getTimeOfRun())
                killProcess(0)
                break
            elif what == "20":
                alg = TKELess(r"TKE", 5)
                alg.run()
                print("Время выполнения :", alg.getTimeOfRun())
                killProcess(0)
                break
            elif what == "3":
                #alg = Decade(r"Decade", 12)
                alg = Decade(r"Decade", 5)
                alg.run()
                print("Время выполнения :", alg.getTimeOfRun())
                killProcess(0)
                break
            elif what == "4":
                alg = NKREKU2(r"NKREKU2", 3)
                alg.run()
                print("Время выполнения :", alg.getTimeOfRun())
                killProcess(0)
                break
            elif what == "5":
                alg = NKREKU3(r"NKREKU3", 3)
                alg.run()
                print("Время выполнения :", alg.getTimeOfRun())
                killProcess(0)
                break
            elif what == "6":
                alg = NKREKU5(r"NKREKU5", 5)
                alg.run()
                print("Время выполнения :", alg.getTimeOfRun())
                killProcess(0)
                break
            elif what == "7":
                alg = NKREKU_PAT(r"NKREKU_PAT(VTV)", 2)
                alg.run()
                print("Время выполнения :", alg.getTimeOfRun())
                killProcess(0)
                break
            elif what == "8":
                alg = NKREKU_PAT_ZBUT_VTV_naselenie(
                                r"NKREKU_PAT_ZBUT(VTV+naselenie)", 2)
                alg.run()
                print("Время выполнения :", alg.getTimeOfRun())
                killProcess(0)
                break
            else:
                print("Неправильный ввод")
        except Exception:
           print("Возникло необработанное исключение")
           print(traceback.format_exc())
    print("Программа завершила работу")
    input()