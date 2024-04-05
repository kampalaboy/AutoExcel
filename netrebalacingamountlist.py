import pandas as pd
import xlwings as xw

# Open the Excel workbook using xlwings
wb = xw.Book(r'C:\Projects\Python\Simba\Data\momoanalysis.xlsx')

# Read the specific sheet using pd.read_excel from pandas
mmsheet = pd.read_excel(wb.fullname, sheet_name='DSD_Agent_Txns_MTD_FEB2023')

# Group the data by 'dsd_line' and find the maximum value in 'Rebalance' column
max_rebalance = mmsheet.groupby('dsd_line')['Rebalance'].max()

# Filter the original data to include only the specified columns
filtered_data = mmsheet[['rebalancing_route', 'sales_region', 'sales_territory', 'parish', 'agent_line']]

# Merge the filtered data with the maximum rebalance values based on 'dsd_line'
merged_data = filtered_data.merge(max_rebalance, on='dsd_line')

# Connect to Excel and open a new workbook or an existing one
wb_output = xw.Book()

# Write the merged data to the Excel workbook
sheet_output = wb_output.sheets['Sheet1']
sheet_output.range('A1').value = merged_data

# Save the workbook
wb_output.save('Output.xlsx')


