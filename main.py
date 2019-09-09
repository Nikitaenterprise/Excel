from FiscalPlan.FiscalPlan import FiscalPlan


if __name__ == "__main__":
    print("I`m main file")
    excelBook = FiscalPlan(r"FiscalPlan\НадходженняНаКР_.xls")
    excelBook.run()