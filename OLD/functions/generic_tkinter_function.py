# Tkinter package is the default python library for file exploration
# Compatible with Windows & Unix, no need to install it through 'pip'
import tkinter.filedialog

# Open a file explorer which prompts the user to select a file (aka the target)
# Params : None
# Return : The target file path
def select_file(file_types):
    path = tkinter.filedialog.askopenfilename(filetypes=file_types)
    return path




