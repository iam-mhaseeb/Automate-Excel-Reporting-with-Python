import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.chart import BarChart, Reference

# Load data
df = pd.read_excel('https://github.com/datagy/pivot_table_pandas/raw/master/sample_pivot.xlsx', parse_dates=['Date'])
df['Date'] = pd.to_datetime(df['Date'])
print(df.head())

# Summary table for East region
filtered = df[df['Region'] == 'East']
quaterly_sales = pd.pivot_table(filtered, index=['Date'], columns='Type', values='Sales', aggfunc='sum')
print('Quaterly Sales Pivot Table')
print(quaterly_sales.head())

# Load data into Excel file
filepath = 'reporting_file.xlsx'
quaterly_sales.to_excel(filepath, sheet_name='Quaterly Sales', startrow=3)

# Making report prettier
wb = load_workbook(filepath)
sheet1 = wb['Quaterly Sales']

sheet1['A1'] = 'Quaterly Sales'
sheet1['A2'] = 'datagy.io'
sheet1['A4'] = 'Quarter'

sheet1['A1'].style = 'Title'
sheet1['A2'].style = 'Headline 2'

for i in range(5, 9):
    sheet1[f'B{i}'].style = 'Currency'
    sheet1[f'C{i}'].style = 'Currency'
    sheet1[f'D{i}'].style = 'Currency'

bar_chart = BarChart()
data = Reference(sheet1, min_col=2, max_col=4, min_row=4, max_row=8)
categories = Reference(sheet1, min_col=1, max_col=1, min_row=5, max_row=8)
bar_chart.add_data(data, titles_from_data=True)
bar_chart.set_categories(categories)
sheet1.add_chart(bar_chart, "F4")

bar_chart.title = 'Sales by Type'
bar_chart.style = 3
wb.save(filename=filepath)
