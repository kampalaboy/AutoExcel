import pandas as pd
import xlwings as xw

# Open the Excel workbook using xlwings
wb = xw.Book(r'C:\Projects\Python\Simba\Data\momoanalysis.xlsx')

# Read the specific sheet using pd.read_excel from pandas
mmsheet = pd.read_excel(wb.fullname, sheet_name='Agents_Served_MTD_FEB2023')

# Create a pivot table using pandas
pivot_table = pd.pivot_table(mmsheet, values=['dsd_active_rate', 'avg_agents_served'], index='dsd_line', columns='sales_territory', aggfunc='sum')

# Filter the pivot table based on conditions
filtered_table = pivot_table[(pivot_table['dsd_active_rate'] == 100) & (pivot_table['avg_agents_served'] > 15)]

# Connect to Excel and open a new workbook or an existing one
wb_output = xw.Book()

# Write the filtered table to the Excel workbook
sheet_output = wb_output.sheets['Sheet1']
sheet_output.range('A1').value = filtered_table

# Count the number of filled rows in each territory column
territory_counts = filtered_table.notnull().sum()

# Create a summary table
summary_table = pd.DataFrame({
    'sales_territory': territory_counts.index,
    'filled_rows_count': territory_counts.values
})

# Color code the territories
summary_table_styled = summary_table.style.background_gradient(cmap='Blues')

# Display the summary table in Excel
sheet_summary = wb_output.sheets.add('Summary')
sheet_summary.range('A1').value = summary_table_styled.data

# Save and close the workbook
wb_output.save('Summary.xlsx')
wb_output.close()
