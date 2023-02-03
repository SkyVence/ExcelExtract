#Library to install | pip pywin32 | pip textfile
from datetime import date 
import textfile
import win32com.client 
from win32com.client import Dispatch
import sys, io

#Request date and time
time = date.fromtimestamp(1326244364)

#Debug Excel
Debug = input("Debug : True/false ?")

# Open up Excel and make it visible
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = Debug

# Ask for file path & Open File
print("Copy File Path (Ex : C:/Users/username/Documents/folderfile)")
file = input("Enter File Path :")
workbook = excel.Workbooks.Open(file)

#Ask User for name of the Sheet
print("What is the name of the sheet you want to extract ?")
name = input("Enter Name : ")

#Ask User for name of the board
print("What is the range of the table you want to extract ? ex : A1:A32")
tableselect = input("Enter range name : ")

#Extract data from range (tablerange)
data_extract = workbook.Worksheets(name).Range(tableselect).value

#Print into console data_extract 
print(data_extract)

# Wait before closing it
_ = input("Press enter to close Excel")
excel.Quit()