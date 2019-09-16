from src.plan import FiscalPlan
from src.tke import TKE

if __name__ == "__main__":
    print("I`m main file")

    #excelBook = FiscalPlan(r"FiscalPlan")
    #excelBook.ppp()
    #excelBook.run()

    tke = TKE(r"TKE")
    tke.run()