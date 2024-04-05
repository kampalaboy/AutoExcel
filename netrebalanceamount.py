import pandas as pd
import xlwings as xw

# Open the Excel workbook using xlwings
wb = xw.Book(r'C:\Projects\Python\Simba\Data\momoanalysis.xlsx')

# Read the specific sheet using pd.read_excel from pandas
mmsheet = pd.read_excel(wb.fullname, sheet_name='DSD_Agent_Txns_MTD_FEB2023')

# Filter the rows where dsd_volume_to_agent_rank is 1
filtered_sheet = mmsheet[mmsheet['dsd_volume_to_agent_rank'] == 1]

# Print the filtered contents of the sheet
print(filtered_sheet)

# Create a pivot table using pandas on the filtered data
pivot_table = pd.pivot_table(filtered_sheet, values=['giving_float', 'receiving_float'], index='dsd_volume_to_agent_rank', aggfunc='sum')

# Print the pivot table
print(pivot_table)

# Open a new workbook using xlwings
wb_output = xw.Book()

# Write the pivot table to the Excel workbook
sheet_output = wb_output.sheets['Sheet1']
sheet_output.range('A1').value = pivot_table

