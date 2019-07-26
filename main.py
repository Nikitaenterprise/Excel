from Excel import TKE

first = TKE("./first.xlsx")
first.start()  
first.hideColumns()
first.save("./out.xlsx")
