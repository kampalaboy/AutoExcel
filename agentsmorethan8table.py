import pandas as pd
import xlwings as xw

# Open the Excel workbook using xlwings
wb = xw.Book(r'C:\Projects\Python\Simba\Data\momoanalysis.xlsx')

# Read the specific sheet using pd.read_excel from pandas
mmsheet = pd.read_excel(wb.fullname, sheet_name='DSD_Agent_Txns_MTD_FEB2023')

# Select the desired columns from the DataFrame
selected_columns = ['agent_line', 'rebalancing_route', 'sales_region', 'sales_territory', 'parish', 'Rebalance']
mmsheet_selected = mmsheet[selected_columns]

# Count the number of rows for each value in 'agent_line' column
row_counts = mmsheet_selected['agent_line'].value_counts()

# Create a DataFrame to store the result
result = pd.DataFrame({'agent_line': row_counts.index, 'row_count': row_counts.values})

# Filter out rows where 'dsd_volume_to_agent_rank' is equal to 1
mmsheet_filtered = mmsheet_selected[mmsheet['dsd_volume_to_agent_rank'] == 1]

# Merge the result DataFrame with the selected columns from the filtered DataFrame
result = pd.merge(result, mmsheet_filtered, on='agent_line')

# Connect to Excel and open a new workbook or an existing one
wb_output = xw.Book()

# Write the result to the Excel workbook
sheet_output = wb_output.sheets['Sheet1']
sheet_output.range('A1').value = result

# Save the workbook
wb_output.save('Output.xlsx')
