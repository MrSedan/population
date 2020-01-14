import openpyxl, wikipedia
wb = openpyxl.load_workbook(filename='./test.xlsx')
sheet = wb[wb.sheetnames[0]]
val = sheet['A1'].value
print(val)