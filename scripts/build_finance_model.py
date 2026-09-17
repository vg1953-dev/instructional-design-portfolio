from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter
from pathlib import Path

OUT = Path('assets/Manufacturing_Cost_Profitability_Analysis.xlsx')
OUT.parent.mkdir(parents=True, exist_ok=True)

wb = Workbook()
ws = wb.active
ws.title = 'Project Overview'

navy = '17324D'
blue = '2F75B5'
light_blue = 'D9EAF7'
green = '70AD47'
light_green = 'E2F0D9'
red = 'C0504D'
light_red = 'FCE4D6'
gray = 'E7E6E6'
dark = '1F1F1F'
white = 'FFFFFF'

thin_gray = Side(style='thin', color='D9D9D9')


def title(ws, text, subtitle=None):
    ws.merge_cells('A1:H1')
    ws['A1'] = text
    ws['A1'].font = Font(size=20, bold=True, color=white)
    ws['A1'].fill = PatternFill('solid', fgColor=navy)
    ws['A1'].alignment = Alignment(vertical='center')
    ws.row_dimensions[1].height = 30
    if subtitle:
        ws.merge_cells('A2:H2')
        ws['A2'] = subtitle
        ws['A2'].font = Font(size=10, italic=True, color='666666')
        ws['A2'].alignment = Alignment(wrap_text=True)
        ws.row_dimensions[2].height = 30


def header_row(ws, row, cols):
    for c, value in enumerate(cols, 1):
        cell = ws.cell(row=row, column=c, value=value)
        cell.font = Font(bold=True, color=white)
        cell.fill = PatternFill('solid', fgColor=blue)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = Border(bottom=thin_gray)


def autosize(ws, max_width=24):
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        width = 10
        for cell in ws[letter]:
            if cell.value is not None:
                width = max(width, min(max_width, len(str(cell.value)) + 2))
        ws.column_dimensions[letter].width = width


def money(cell):
    cell.number_format = '$#,##0;[Red]-$#,##0'


def money1(cell):
    cell.number_format = '$#,##0.0,,"M";[Red]-$#,##0.0,,"M"'


def pct(cell):
    cell.number_format = '0.0%'

# 1. PROJECT OVERVIEW
ws = wb['Project Overview']
title(ws, 'Manufacturing Cost & Profitability Analysis',
      'Self-directed finance & analytics case study using illustrative manufacturing data. No company or Tesla data is used.')
ws['A4'] = 'Business Question'
ws['A4'].font = Font(bold=True, color=navy, size=12)
ws.merge_cells('A5:H6')
ws['A5'] = ('How do changes in production volume, selling price, material cost, labor, variable overhead, and fixed costs '
            'affect manufacturing profitability—and which drivers deserve the most management attention?')
ws['A5'].alignment = Alignment(wrap_text=True, vertical='top')

summary = [
    ('Base-case Revenue', "='Unit Economics'!B11"),
    ('Operating Profit', "='Unit Economics'!B15"),
    ('Operating Margin', "='Unit Economics'!B16"),
    ('Break-even Units', "='Unit Economics'!B17"),
]
for i, (label, formula) in enumerate(summary, 0):
    col = 1 + i * 2
    ws.cell(8, col, label).font = Font(bold=True, color='666666')
    ws.cell(9, col, formula).font = Font(size=18, bold=True, color=navy)
    ws.merge_cells(start_row=8, start_column=col, end_row=8, end_column=col+1)
    ws.merge_cells(start_row=9, start_column=col, end_row=10, end_column=col+1)
    ws.cell(9, col).alignment = Alignment(vertical='center')
    ws.cell(8, col).fill = PatternFill('solid', fgColor=gray)
    ws.cell(9, col).fill = PatternFill('solid', fgColor='F7F9FB')

money1(ws['A9']); money1(ws['C9']); pct(ws['E9']); ws['G9'].number_format = '#,##0'

ws['A12'] = 'Model Components'
ws['A12'].font = Font(bold=True, color=navy, size=12)
components = [
    ('Assumptions', 'Driver inputs for volume, price, unit costs, and fixed costs.'),
    ('Unit Economics', 'Contribution margin, break-even volume, operating profit, and margin of safety.'),
    ('Monthly Forecast', 'Budget vs. illustrative actual performance and variance analysis.'),
    ('Scenario Analysis', 'Downside, base, upside, and sensitivity analysis.'),
    ('Dashboards', 'Executive, variance, and scenario views for management communication.'),
]
for r, (a, b) in enumerate(components, 13):
    ws[f'A{r}'] = a; ws[f'A{r}'].font = Font(bold=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    ws.cell(r, 2, b)

for c in range(1, 9): ws.column_dimensions[get_column_letter(c)].width = 16
ws.freeze_panes = 'A4'

# 2. ASSUMPTIONS
ws = wb.create_sheet('Assumptions')
title(ws, 'Assumptions', 'Editable driver inputs for the base manufacturing case.')
header_row(ws, 4, ['Driver', 'Base Input', 'Unit / Note'])
assumptions = [
    ('Annual production / sales volume', 120000, 'units'),
    ('Selling price per unit', 420, '$ / unit'),
    ('Material cost per unit', 210, '$ / unit'),
    ('Direct labor per unit', 48, '$ / unit'),
    ('Variable overhead per unit', 25, '$ / unit'),
    ('Variable selling cost per unit', 10, '$ / unit'),
    ('Annual fixed manufacturing costs', 8500000, '$ / year'),
    ('Annual fixed SG&A / other costs', 4100000, '$ / year'),
]
for r, row in enumerate(assumptions, 5):
    for c, val in enumerate(row, 1): ws.cell(r, c, val)
    ws.cell(r, 2).fill = PatternFill('solid', fgColor=light_blue)
    if '$' in str(row[2]): money(ws.cell(r, 2))
ws['A15'] = 'Illustrative portfolio assumptions only — not sourced from Tesla or any employer.'
ws['A15'].font = Font(italic=True, color='666666')
ws.merge_cells('A15:F15')
autosize(ws, 34)
ws.freeze_panes = 'A5'

# 3. UNIT ECONOMICS
ws = wb.create_sheet('Unit Economics')
title(ws, 'Unit Economics', 'Base-case manufacturing profitability and break-even economics.')
header_row(ws, 4, ['Metric', 'Value', 'Formula / Interpretation'])
rows = [
    ('Selling price / unit', "='Assumptions'!B6", 'Revenue per unit'),
    ('Material cost / unit', "='Assumptions'!B7", 'Variable cost'),
    ('Direct labor / unit', "='Assumptions'!B8", 'Variable cost'),
    ('Variable overhead / unit', "='Assumptions'!B9", 'Variable cost'),
    ('Variable selling / unit', "='Assumptions'!B10", 'Variable cost'),
    ('Total variable cost / unit', '=SUM(B6:B9)', 'Sum of variable costs'),
    ('Contribution / unit', '=B5-B10', 'Price less variable cost'),
    ('Revenue', "='Assumptions'!B5*B5", 'Annual volume × selling price'),
    ('Total variable cost', "='Assumptions'!B5*B10", 'Annual volume × variable cost / unit'),
    ('Contribution margin', '=B11-B12', 'Revenue less total variable cost'),
    ('Total fixed costs', "='Assumptions'!B11+'Assumptions'!B12", 'Fixed manufacturing + fixed SG&A'),
    ('Operating profit', '=B13-B14', 'Contribution margin less fixed costs'),
    ('Operating margin', '=B15/B11', 'Operating profit ÷ revenue'),
    ('Break-even units', '=B14/B10', 'Fixed costs ÷ contribution / unit'),
    ('Margin of safety', "=('Assumptions'!B5-B17)/'Assumptions'!B5", 'Excess volume above break-even'),
]
for r, (metric, formula, note) in enumerate(rows, 5):
    ws.cell(r, 1, metric); ws.cell(r, 2, formula); ws.cell(r, 3, note)
    if r in [10, 11, 13, 15, 17]: ws.cell(r, 1).font = Font(bold=True, color=navy)
for r in range(5, 18):
    if r in [16, 18]: pct(ws.cell(r, 2))
    elif r == 17: ws.cell(r, 2).number_format = '#,##0'
    else: money(ws.cell(r, 2))
autosize(ws, 38)
ws.freeze_panes = 'A5'

# 4. MONTHLY FORECAST / VARIANCE
ws = wb.create_sheet('Monthly Forecast')
title(ws, 'Monthly Forecast', 'Budget vs. illustrative actual results and operating-profit variance.')
headers = ['Month','Budget Units','Actual Units','Budget Price','Actual Price','Budget Material / Unit','Actual Material / Unit','Budget Revenue','Actual Revenue','Budget Op Profit','Actual Op Profit','Profit Variance']
header_row(ws, 4, headers)
months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
budget_units = [8500,9000,9500,9500,10000,10000,10000,10500,10500,10500,10000,10000]
actual_units = [8200,8700,9300,8800,9500,9000,8600,10200,10500,11100,10800,11100]
actual_price = [417,418,419,416,417,415,414,418,419,421,422,422]
actual_mat = [214,215,214,216,218,219,221,218,216,214,213,212]
for i, m in enumerate(months, 5):
    j=i-5
    ws.cell(i,1,m); ws.cell(i,2,budget_units[j]); ws.cell(i,3,actual_units[j])
    ws.cell(i,4,"='Assumptions'!B6"); ws.cell(i,5,actual_price[j])
    ws.cell(i,6,"='Assumptions'!B7"); ws.cell(i,7,actual_mat[j])
    ws.cell(i,8,f'=B{i}*D{i}'); ws.cell(i,9,f'=C{i}*E{i}')
    # Allocate fixed costs monthly; actual uses actual price/material, same labor & overhead assumptions
    ws.cell(i,10,f'=B{i}*(D{i}-(F{i}+\'Assumptions\'!B8+\'Assumptions\'!B9+\'Assumptions\'!B10))-(\'Assumptions\'!B11+\'Assumptions\'!B12)/12')
    ws.cell(i,11,f'=C{i}*(E{i}-(G{i}+\'Assumptions\'!B8+\'Assumptions\'!B9+\'Assumptions\'!B10))-(\'Assumptions\'!B11+\'Assumptions\'!B12)/12')
    ws.cell(i,12,f'=K{i}-J{i}')
for c in range(4,13):
    for r in range(5,17): money(ws.cell(r,c))
ws.cell(18,1,'TOTAL'); ws.cell(18,1).font=Font(bold=True)
for c in [2,3,8,9,10,11,12]:
    letter=get_column_letter(c); ws.cell(18,c,f'=SUM({letter}5:{letter}16)'); ws.cell(18,c).font=Font(bold=True)
for c in [8,9,10,11,12]: money(ws.cell(18,c))
autosize(ws, 22); ws.freeze_panes='A5'

# 5. SCENARIO ANALYSIS
ws = wb.create_sheet('Scenario Analysis')
title(ws, 'Scenario Analysis', 'Downside, base, upside, and individual profitability sensitivities.')
header_row(ws, 4, ['Driver','Downside','Base','Upside'])
scenario_rows = [
    ('Volume', 100000, 120000, 135000),
    ('Price / unit', 400, 420, 430),
    ('Material / unit', 225, 210, 202),
    ('Labor / unit', 52, 48, 47),
    ('Variable OH / unit', 28, 25, 24),
    ('Variable selling / unit', 10, 10, 10),
    ('Fixed costs', 13000000, 12600000, 12350000),
]
for r,row in enumerate(scenario_rows,5):
    for c,val in enumerate(row,1): ws.cell(r,c,val)
for c in range(2,5):
    for r in range(6,12): money(ws.cell(r,c))
header_row(ws, 13, ['Metric','Downside','Base','Upside'])
for c in range(2,5):
    col=get_column_letter(c)
    ws.cell(14,c,f'={col}5*{col}6')
    ws.cell(15,c,f'={col}7+{col}8+{col}9+{col}10')
    ws.cell(16,c,f'={col}5*({col}6-{col}15)-{col}11')
    ws.cell(17,c,f'={col}16/{col}14')
    ws.cell(18,c,f'={col}11/({col}6-{col}15)')
for r,label in enumerate(['Revenue','Variable cost / unit','Operating profit','Operating margin','Break-even units'],14): ws.cell(r,1,label)
for c in range(2,5):
    money(ws.cell(14,c)); money(ws.cell(15,c)); money(ws.cell(16,c)); pct(ws.cell(17,c)); ws.cell(18,c).number_format='#,##0'

header_row(ws, 21, ['Sensitivity Lever','Change','Annual Op Profit Impact'])
sens = [
    ('Volume', '+10,000 units', '=10000*(\'Unit Economics\'!B10)'),
    ('Material cost', '-$5 / unit', "='Assumptions'!B5*5"),
    ('Selling price', '+1%', "='Assumptions'!B5*'Assumptions'!B6*1%"),
    ('Fixed costs', '-$500,000', '=500000'),
]
for r,row in enumerate(sens,22):
    for c,val in enumerate(row,1): ws.cell(r,c,val)
    money(ws.cell(r,3))
autosize(ws, 26); ws.freeze_panes='A5'

# 6. EXECUTIVE DASHBOARD
ws = wb.create_sheet('Dashboard')
title(ws, 'Executive Performance Dashboard', 'Management view of profitability, break-even economics, scenarios, and key sensitivities.')
# KPI cards
kpis = [('Revenue',"='Unit Economics'!B11"),('Operating Profit',"='Unit Economics'!B15"),('Operating Margin',"='Unit Economics'!B16"),('Break-even Units',"='Unit Economics'!B17")]
for i,(lab,form) in enumerate(kpis):
    c=1+i*2
    ws.merge_cells(start_row=4,start_column=c,end_row=4,end_column=c+1)
    ws.merge_cells(start_row=5,start_column=c,end_row=6,end_column=c+1)
    ws.cell(4,c,lab).fill=PatternFill('solid',fgColor=gray); ws.cell(4,c).font=Font(bold=True,color='666666')
    ws.cell(5,c,form).fill=PatternFill('solid',fgColor='F7F9FB'); ws.cell(5,c).font=Font(size=17,bold=True,color=navy)
    ws.cell(5,c).alignment=Alignment(vertical='center')
money1(ws['A5']); money1(ws['C5']); pct(ws['E5']); ws['G5'].number_format='#,##0'
# scenario support table
ws['A9']='Scenario'; ws['B9']='Operating Profit'; ws['D9']='Scenario'; ws['E9']='Break-even Units'
for cell in ['A9','B9','D9','E9']: ws[cell].font=Font(bold=True,color=white); ws[cell].fill=PatternFill('solid',fgColor=blue)
for idx, (sc, col) in enumerate([('Downside','B'),('Base','C'),('Upside','D')],10):
    ws.cell(idx,1,sc); ws.cell(idx,2,f"='Scenario Analysis'!{col}16"); money(ws.cell(idx,2))
    ws.cell(idx,4,sc); ws.cell(idx,5,f"='Scenario Analysis'!{col}18"); ws.cell(idx,5).number_format='#,##0'
chart=BarChart(); chart.type='col'; chart.style=10; chart.title='Scenario Operating Profit'; chart.y_axis.title='Operating Profit ($)'; chart.height=8; chart.width=11
chart.add_data(Reference(ws,min_col=2,min_row=9,max_row=12),titles_from_data=True); chart.set_categories(Reference(ws,min_col=1,min_row=10,max_row=12)); chart.legend=None; chart.dLbls=DataLabelList(); chart.dLbls.showVal=True
ws.add_chart(chart,'A14')
chart2=BarChart(); chart2.type='col'; chart2.style=10; chart2.title='Break-even Units'; chart2.height=8; chart2.width=11
chart2.add_data(Reference(ws,min_col=5,min_row=9,max_row=12),titles_from_data=True); chart2.set_categories(Reference(ws,min_col=4,min_row=10,max_row=12)); chart2.legend=None; chart2.dLbls=DataLabelList(); chart2.dLbls.showVal=True
ws.add_chart(chart2,'E14')
for c in range(1,9): ws.column_dimensions[get_column_letter(c)].width=16
ws.sheet_view.showGridLines=False

# 7. VARIANCE DASHBOARD
ws = wb.create_sheet('Variance Dashboard')
title(ws, 'Budget vs. Actual Variance Dashboard', 'Monthly performance view highlighting revenue and operating-profit divergence from plan.')
# KPI block
labels = [('Budget Op Profit',"='Monthly Forecast'!J18"),('Actual Op Profit',"='Monthly Forecast'!K18"),('Profit Variance',"='Monthly Forecast'!L18"),('Actual Units',"='Monthly Forecast'!C18")]
for i,(lab,form) in enumerate(labels):
    c=1+i*2
    ws.merge_cells(start_row=4,start_column=c,end_row=4,end_column=c+1); ws.merge_cells(start_row=5,start_column=c,end_row=6,end_column=c+1)
    ws.cell(4,c,lab).fill=PatternFill('solid',fgColor=gray); ws.cell(4,c).font=Font(bold=True,color='666666')
    ws.cell(5,c,form).fill=PatternFill('solid',fgColor='F7F9FB'); ws.cell(5,c).font=Font(size=17,bold=True,color=navy)
for c in ['A5','C5','E5']: money1(ws[c])
ws['G5'].number_format='#,##0'
# hidden-ish support table
header_row(ws,9,['Month','Budget Op Profit','Actual Op Profit'])
for r in range(10,22):
    src=r-5
    ws.cell(r,1,f"='Monthly Forecast'!A{src}")
    ws.cell(r,2,f"='Monthly Forecast'!J{src}")
    ws.cell(r,3,f"='Monthly Forecast'!K{src}")
    money(ws.cell(r,2)); money(ws.cell(r,3))
line=LineChart(); line.style=13; line.title='Monthly Operating Profit — Budget vs. Actual'; line.y_axis.title='Operating Profit ($)'; line.height=10; line.width=18
line.add_data(Reference(ws,min_col=2,max_col=3,min_row=9,max_row=21),titles_from_data=True); line.set_categories(Reference(ws,min_col=1,min_row=10,max_row=21)); line.legend.position='b'
ws.add_chart(line,'A24')
for c in range(1,9): ws.column_dimensions[get_column_letter(c)].width=16
ws.sheet_view.showGridLines=False

# 8. SCENARIO DASHBOARD
ws = wb.create_sheet('Scenario Dashboard')
title(ws, 'Scenario & Sensitivity Dashboard', 'Decision-support view of the operating levers with the largest modeled profit impact.')
header_row(ws,4,['Sensitivity Lever','Change','Annual Op Profit Impact'])
for r in range(5,9):
    src=r+17
    ws.cell(r,1,f"='Scenario Analysis'!A{src}")
    ws.cell(r,2,f"='Scenario Analysis'!B{src}")
    ws.cell(r,3,f"='Scenario Analysis'!C{src}")
    money(ws.cell(r,3))
bar=BarChart(); bar.type='bar'; bar.style=10; bar.title='Annual Operating-Profit Sensitivity'; bar.x_axis.title='Profit Impact ($)'; bar.height=9; bar.width=16
bar.add_data(Reference(ws,min_col=3,min_row=4,max_row=8),titles_from_data=True); bar.set_categories(Reference(ws,min_col=1,min_row=5,max_row=8)); bar.legend=None; bar.dLbls=DataLabelList(); bar.dLbls.showVal=True
ws.add_chart(bar,'A11')
ws['A30']='Portfolio note: All values are illustrative case-study data and are not Tesla or employer data.'
ws['A30'].font=Font(italic=True,color='666666'); ws.merge_cells('A30:H30')
for c in range(1,9): ws.column_dimensions[get_column_letter(c)].width=18
ws.sheet_view.showGridLines=False

# workbook properties and save
wb.calculation.fullCalcOnLoad = True
wb.calculation.forceFullCalc = True
wb.calculation.calcMode = 'auto'
wb.save(OUT)
print(f'Created {OUT} ({OUT.stat().st_size} bytes)')
