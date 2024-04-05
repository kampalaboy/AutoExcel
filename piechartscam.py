import pandas as pd
import xlwings as xw
import matplotlib.pyplot as plt

# Open the Excel workbook using xlwings
wb = xw.Book(r'C:\Projects\Python\Simba\Data\momoanalysis.xlsx')

# Read the specific sheet using pd.read_excel from pandas
mmsheet = pd.read_excel(wb.fullname, sheet_name='DSD_Agent_Txns_MTD_FEB2023')

# Count the number of rows for each value in 'agent_line' column
row_counts = mmsheet['agent_line'].value_counts()

# Create a DataFrame to store the result
result = pd.DataFrame({'agent_line': row_counts.index, 'row_count': row_counts.values})

# Calculate the percentage of each row count
result['percentage'] = (result['row_count'] / result['row_count'].sum()) * 100

# Connect to Excel and open a new workbook or an existing one
wb_output = xw.Book()

# Write the result to the Excel workbook
sheet_output = wb_output.sheets['Sheet1']
sheet_output.range('A1').value = result

# Create a pie chart
plt.figure(figsize=(6, 6))
plt.pie(result['row_count'], labels=result['agent_line'], autopct='%1.1f%%')
plt.title('Agent Line Row Counts')

# Save the chart to an image file
chart_path = 'pie_chart.png'
plt.savefig(chart_path)

# Insert the chart image into the Excel workbook
sheet_output.pictures.add(chart_path, name='PieChart', update=True)

# Save the workbook
wb_output.save('Output.xlsx')


