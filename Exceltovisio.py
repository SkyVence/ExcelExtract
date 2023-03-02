import win32com.client as win32
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Create GUI window
root = tk.Tk()
root.withdraw()

# Ask user to select Excel file
excel_file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))

# Ask user to select output file path
output_file_path = filedialog.asksaveasfilename(title="Save Visio File As", defaultextension=".vsd", filetypes=(("Visio Files", "*.vsd"), ("All Files", "*.*")))

# Load data from Excel file
df = pd.read_excel(excel_file_path)

# Create Visio application object
visio = win32.gencache.EnsureDispatch('Visio.Application')

# Create new document
doc = visio.Documents.Add()

# Create a new page
page = doc.Pages.Add()

# Add shapes to the page
for i, row in df.iterrows():
    shape = page.DrawRectangle(row['x'], row['y'], row['width'], row['height'])
    shape.Text = row['text']

# Save the document
doc.SaveAs(output_file_path)

# Close the document and Visio application
doc.Close()
visio.Quit()
