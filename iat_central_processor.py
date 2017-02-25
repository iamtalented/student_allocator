from openpyxl import load_workbook, Workbook 

wb = load_workbook("Consolidated Preferences (Registration and Attendance Lists).xlsx")

for sheet in wb.get_sheet_names():
    