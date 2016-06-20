import openpyxl
wb = openpyxl.load_workbook('nat_ground_contacts.xls')
sheet = wb.get_sheet_by_name('REFINED')
cellA = sheet['A1'].value

print cellA
