import pandas as pd
import xlwings as xw

# Open the Excel workbook using xlwings
wb = xw.Book(r'C:\Projects\Python\Simba\Data\momoanalysis.xlsx')

# Read the specific sheet using pd.read_excel from pandas
mmsheet = pd.read_excel(wb.fullname, sheet_name='Agents_Served_MTD_FEB2023')

# Print the contents of the sheet
print(mmsheet)

# Create a pivot table using pandas
pivot_table = pd.pivot_table(mmsheet, values=['dsd_active_rate', 'avg_agents_served'], index='dsd_line', columns='sales_territory', aggfunc='sum')

# Connect to Excel and open a new workbook or an existing one
wb_output = xw.Book()

# Write the pivot table to the Excel workbook
sheet_output = wb_output.sheets['Sheet1']
sheet_output.range('A1').value = pivot_table

