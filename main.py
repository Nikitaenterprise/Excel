from src.plan import FiscalPlan


if __name__ == "__main__":
    print("I`m main file")

    excelBook = FiscalPlan(r"FiscalPlan")
    #excelBook.ppp()
    excelBook.run()

    #tke = TKE()