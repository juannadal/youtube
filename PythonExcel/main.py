from openpyxl import Workbook
from openpyxl import load_workbook

'''
wb = Workbook()

ws = wb.active


ws1 = wb.create_sheet("Mysheet")

ws.title = "New Title"


wb.save('balances.xlsx')


'''
wb2 = load_workbook('balances.xlsx')

myworksheet = wb2.get_sheet_by_name("Mysheet")

myworksheet["A2"].value = "Juan Nadal"


wb2.save('balances.xlsx')