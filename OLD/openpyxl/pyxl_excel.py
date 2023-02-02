#DOC = https://openpyxl.readthedocs.io/en/stable/tutorial.html#playing-with-data

from openpyxl import Workbook 
from openpyxl import load_workbook


#Create a file 
workbook = Workbook()
sheet = workbook.active

#Add Text to cell
sheet["A1"] = "New Cell"

#Adding ROW
sheet.insert_rows(7)

#Adding Columns

  
#Save-File
workbook.save(filename="hello_world.xlsx")

#If you want to load file : Import Load_workbook | You can also set a priviledge : read_only (True/False). If you want only the result from formulas then replace/add data_only
#workbook = load_workbook(filename="loading_world.xls", read_only=True/false)
