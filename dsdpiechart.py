import pandas as pd
import xlwings as xw

# Open the Excel workbook using xlwings
wb = xw.Book(r'C:\Projects\Python\Simba\Data\momoanalysis.xlsx')

# Read the specific sheet using pd.read_excel from pandas
mmsheet = pd.read_excel(wb.fullname, sheet_name='DSD_Agent_Txns_MTD_FEB2023')

# Create a pivot table using pandas
pivot_table = pd.pivot_table(mmsheet, values='dsd_line', index='sales_territory', aggfunc='count')

# Connect to Excel and open a new workbook or an existing one
wb_output = xw.Book()

# Write the pivot table to the Excel workbook
sheet_output = wb_output.sheets['Sheet1']
sheet_output.range('A1').value = pivot_table

# Create a pie chart
chart = sheet_output.charts.add()
chart.set_source_data(sheet_output.range('A1').expand())

# Set the chart type to a pie chart
chart.chart_type = 'pie'

# Apply color coding to the chart segments by sales territory
chart.series_collection(1).format.fill.visible = True
chart.series_collection(1).format.fill.solid()
chart.series_collection(1).format.fill.fore_color.rgb = 255  # Set the color to red
chart.series_collection(2).format.fill.visible = True
chart.series_collection(2).format.fill.solid()
chart.series_collection(2).format.fill.fore_color.rgb = 65535  # Set the color to yellow
# Repeat the above lines for each sales territory and color you want to assign

# Optional: Adjust chart properties
chart.has_legend = True  # Show legend
chart.legend.position = 'right'  # Position the legend to the right of the chart
chart.chart_title.text = 'Sales Territories'  # Set the chart title

# Adjust the chart position and size
chart.left = 300
chart.top = 100
chart.width = 400
chart.height = 300

# Add data labels to the chart segments
chart.data_labels.number_format = '0'
chart.data_labels.position = 'outside_end'
chart.data_labels.show_category = True
chart.data_labels.show_value = True
