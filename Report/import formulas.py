import formulas


fpath = r"C:\Users\frick\Documents\VT1/Fussballteam.xlsx"
dir_output = "output"

xl_model = formulas.ExcelModel().loads(fpath).finish()
xl_model.calculate(
     inputs={
         "'[Fussballteam.xlsx]Arbeit'!A27": "new"  # To impose a value to B3 cell.
     })

xl_model.write(dirpath=dir_output)




#'https://d.docs.live.net/149fa13a5026133a/Fussballteam.xlsx'