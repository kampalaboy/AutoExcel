import pandas as pd
import xlwings as xw

# Open the Excel workbook using xlwings
wb = xw.Book(r'C:\Projects\Python\Simba\Data\exercise.xlsx')

# Read the specific sheet using pd.read_excel from pandas
mmsheet = pd.read_excel(wb.fullname, sheet_name='RebalanceRanked1')

# Count the number of rows for each value in 'agent_line' column
row_counts = mmsheet['agent_line'].value_counts()

# Create a DataFrame to store the result
result = pd.DataFrame({'agent_line': row_counts.index, 'row_count': row_counts.values})

# Connect to Excel and open a new workbook or an existing one
wb_output = xw.Book()

# Write the result to the Excel workbook
sheet_output = wb_output.sheets['Sheet1']
sheet_output.range('A1').value = result

# Save the workbook
wb_output.save('Output.xlsx')


