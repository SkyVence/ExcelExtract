# Win32com package (known as 'pywin32') provides access to Windows APIs
# Manipulate MS Office Apps, require installation 'pip install pywin32'
import win32com.client


#->->->->->->->-> WORKBOOK & SHEET RELATED 

# Open an excel workbook
# Params : app_ - An opened excel application (OBJECT)
#          filepath - The workbook filepath (STRING)
# Return : The opened workbook (OBJECT)
def open_workbook(app_, filepath):
    return app_.Workbooks.Open(filepath)

# Enumerate the sheets in an open excel document
# Params : file_ - An opened excel file (OBJECT)
# Return : List of sheets (LIST)
def get_sheets_name(file_):
    sheets_obj = file_.Sheets
    sheet_namelist = [sheet.Name for sheet in sheets_obj]
    return sheet_namelist

# Change the active sheet in an open excel document
# Params : file_ - An opened excel file (OBJECT)
#          sheet_name - The excel sheet name (STRING)
# Return : None
def set_active_sheet(file_, sheet_name):
    file_.Worksheets(sheet_name).Activate()


#->->->->->->->-> STRUCTURE RELATED 

# Change the cell background color in a specific sheet
# Params : file_ - An opened excel file (OBJECT)
#          sheet_name - The excel sheet name (STRING)
#          cell_code - The excel cell location "(col, line)"
#            ├─ (INT,INT) ex:(27,1) OR (STR,INT) ex:("AA",1)
#            └─ Must be uppercase if cols are passed as str.
#          colour - The color in HEX format (STRING)
# Return : None
def set_cell_color(file_, sheet_name, cell_, color):
    col, line = cell_[0],cell_[1]
    if type(col) != int:
        col = convert_col_letter_to_idx(col)
    file_.sheets(sheet_name).Cells(line, col).Interior.Color = color

# Return the cell background color in a specific sheet
# Params : file_ - An opened excel file (OBJECT)
#          sheet_name - The excel sheet name (STRING)
#          cell_code - The excel cell location "(col, line)"
#            ├─ (INT,INT) ex:(27,1) OR (STR,INT) ex:("AA",1)
#            └─ Must be uppercase if cols are passed as str.
# Return : The HEX color code (STRING)
def get_cell_color(file_, sheet_name, cell_):
    col, line = cell_[0],cell_[1]
    if type(col) != int:
        col = convert_col_letter_to_idx(col)
    return file_.sheets(sheet_name).Cells(line, col).Interior.Color


#->->->->->->->-> CONTENT RELATED 

# Return the content of all used cell (range) in a specific sheet
# This will form a table from the first filled cell (top left)
# To the last filled cell (bottom right), unused cell = none
# Params : file_ - An opened excel file (OBJECT)
#          sheet_name - The excel sheet name (STRING)
# Return : The sheet content (STRING)
def get_sheet_content(file_, sheet_name):
    return file_.sheets(sheet_name).UsedRange

# Return the content of all named cell range in a specific sheet
# Params : file_ - An opened excel file (OBJECT)
#          sheet_name - The excel sheet name (STRING)
#          named_range - A cell named range (STRING)
# Return : The range content (TUPLES)
def get_named_range_content(file_, sheet_name, named_range):
    return file_.sheets(sheet_name).Range(named_range).Value

# Return the content of a named table in a specific sheet
# Params : file_ - An opened excel file (OBJECT)
#          sheet_name - The excel sheet name (STRING)
#          named_range - A cell named range (STRING)
# Return : The table data, and headers (TUPLES, TUPLES)
def get_named_table_content(file_, sheet_name, named_range):
    table_ = file_.sheets(sheet_name).ListObjects(named_range)
    return table_.DataBodyRange.Value , table_.HeaderRowRange.Value
    


#->->->->->->->-> MISC HELPER 


# Convert an Excel Col Name to its Index
# Params : col_id - A col name (STRING)
# Return : The col index (INT)
def convert_col_letter_to_idx(col_id):
    tmp = 0
    #Loop through each letter forming the col_id (ex: for "AB" there is two iteration)
    for i in range(len(col_id)):
        #Multiply 'current 'position' by alphabet length (starting at 0)
        tmp *= 26
        #Get the first alphabet UPPERCASE index (A=65)
        capital_A_idx = ord('A')
        #Get the current alphabet UPPERCASE index (ex:G -> G=71)
        current_capital_letter_idx = ord(col_id[i]) 
        #Soustract current to first letter idx to get position, add +1 for displacement
        tmp += current_capital_letter_idx - capital_A_idx + 1 
    return tmp