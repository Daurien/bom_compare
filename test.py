
import openpyxl

# Load the workbook
file_path = 'Numero Spec - Copie.xlsm'
workbook = openpyxl.load_workbook(file_path, keep_vba=True)

# Select the active sheet or specify the sheet name
sheet = workbook.active  # or workbook['SheetName'] if you know the sheet name

# Modify cell E4
sheet.cell(row=5, column=5, value='TEST SAMBLESDS')

# Save the workbook
workbook.save(file_path)
