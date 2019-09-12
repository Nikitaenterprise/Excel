from FiscalPlan.FiscalPlan import FiscalPlan


if __name__ == "__main__":
    print("I`m main file")
    excelBook = FiscalPlan(r"FiscalPlan\09_Прогнозне надходження коштів у вересні 2019 року.xlsx",
                            r"FiscalPlan\18.xls",
                            r"FiscalPlan\Всі Категорії. (без спожив. за СЕРПЕНЬ)_ПАТ.xls",
                            r"FiscalPlan\Всі Категорії. (без спожив. за СЕРПЕНЬ)_ЗБУТ.xls",
                            r"FiscalPlan\ТЕЦ.xls"
                            )
    excelBook.run()
    excelBook.deleteFiles()
