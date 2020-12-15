from openpyxl import load_workbook


sealevel = load_workbook('CTWP_Sea_Level_Report_November2020.xlsx')
print(sealevel.active['A10'].value)
