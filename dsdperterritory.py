import pandas as pd
import xlwings as xw

# Open the Excel workbook using xlwings
wb = xw.Book(r'C:\Projects\Python\Simba\Data\momoanalysis.xlsx')

# Read the specific sheet using pd.read_excel from pandas
mmsheet = pd.read_excel(wb.fullname, sheet_name='DSDsAnalysisTargetAchievement')

# Print the contents of the sheet
print(mmsheet)

pivot_table = pd.pivot_table(mmsheet, values='dsd_line', index='sales_territory', aggfunc='count')

# Rename the column to 'dsd_line_count' for clarity
pivot_table = pivot_table.rename(columns={'dsd_line': 'dsd_line_count'})

# Print the pivot table
print(pivot_table)

wb_output = xw.Book()

# Write the filtered table to the Excel workbook
sheet_output = wb_output.sheets['Sheet1']
sheet_output.range('A1').value = pivot_table

