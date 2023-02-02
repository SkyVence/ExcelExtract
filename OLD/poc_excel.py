activate_debug = True

def debug_pause(print_=""):
    if activate_debug:
        print(print_)
        input("<- Press Enter To Continue ->")


###################################################################
######                   SETUP & FUNCTION                    ######
###################################################################

#->->->->->->->->->   LOCAL LIBRARY IMPORTS   <-<-<-<-<-<-<-<-<-<-#

# OS and SYS package provides APIs to manipulate the operating system
# Compatible with Windows & Unix, no need to install it through 'pip'
import os,sys

# OS and SYS are used to inject the function subfolder required by this script
# This won't affect the actual system path variables only this script instance

# Firt find the script location (path) and add the (functions) subfolder to it
function_subfolder = os.path.dirname(os.path.realpath(__file__)) + '\\functions'
# Then append that "sub-path" to the list of path known by the python instance
sys.path.append(function_subfolder)

# Pandas package is a fast, flexible and powerfull data structure manipulation
# Compatible with Windows & Unix, require installation 'pip install pandas'
import pandas as pandas_


#->->->->->->->->    EXTERNAL LIBRARY IMPORTS   <-<-<-<-<-<-<-<-<-#

# Win32com package (known as 'pywin32') provides access to Windows APIs
# Manipulate MS Office Apps, require installation 'pip install pywin32'
# import win32com.client (import is done in functions external files)

# Tkinter package is the default python library for file exploration
# Compatible with Windows & Unix, no need to install it through 'pip'
# import tkinter.filedialog (import is done in functions ext. files)


#->->->->->->->->->     FUNCTION DEFINITION   <-<-<-<-<-<-<-<-<-<-#

# Functions have been extracted to other files for maintainability
# All those files are located in the subfolder called "functions"
# That subfolder MUST be located at the same level of THIS file !
#   /root-folder
#    ├─ /functions
#    │   ├── other-function-file.py
#    │   └── some-function-file.py
#    └── this-script-file.py

import generic_tkinter_function as explorer_
import generic_win32_function as win32_
import excel_win32_function as excel_



###################################################################
######                      MAIN CODE                        ######
###################################################################

#->->->->->->->->     EXCEL PART - LOAD DATA    <-<-<-<-<-<-<-<-<-#
#Start Excel
excel_app = win32_.start_app(application_codename = "Excel.Application",background = (not activate_debug))

#Get Workbook FilePath
# > > > > > > > excel_filepath = explorer_.select_file(file_types=[("Excel files", ".xlsx .xls")])

#Open Workbook
# > > > > > > > excel_file = open_workbook(excel_app, excel_filepath)
# > > > > > > > debug_pause("Opened excel workbook from : " + excel_filepath)
excel_file = excel_.open_workbook(excel_app, "c:/Users/antoine.mathie/Documents/PY_Project/data_test.xlsx")

#Retrieve the first (main) sheet name
main_sheet_name = excel_.get_sheets_name(excel_file)[0]

#debug_pause(excel_.set_active_sheet(excel_file,main_sheet_name))
#debug_pause(excel_.get_sheet_content(excel_file,main_sheet_name))
#debug_pause(excel_.get_cell_color(excel_file,main_sheet_name,("A",2)))
#debug_pause(excel_.set_cell_color(excel_file,main_sheet_name,("A",2),"800080"))

main_data, main_header = excel_.get_named_table_content(excel_file,main_sheet_name,"TABLE")
main_table = pandas_.DataFrame(main_data,columns=main_header[0])

debug_pause(main_table)

for idx in main_table.index:
    print(main_table.loc[[idx]])
    print("______")

debug_pause()

#VISIO SIMPLE SHAPE WITH NAME IP
#THEN EXCEL READ STACK (HOSTNAME + #MEMBER) in an OBJECt
#MAKE A VISIO STK SHAPES
#THEN READ CONNECTORS
#THEN DRAW CONNECTORS

win32_.save_doc_and_quit_app(excel_app, excel_file)
