import openpyxl
import cmccall
from openpyxl.styles import Font
from tkinter.filedialog import asksaveasfilename

wb = openpyxl.load_workbook('templatexlsx.xlsx')
yesFont = Font(color='00FF00', bold='True')
noFont = Font(color='FF0000')

one_to_ten = wb.worksheets[0]
ten_to_twenty = wb.worksheets[1]
twenty_to_fifty = wb.worksheets[2]

# Fill Excel
print("Now running: $1-10M...")
one_to_ten_total = cmccall.handle(1, 10, one_to_ten)
print("Now running: $10-20M...")
ten_to_twenty_total = cmccall.handle(10, 20, ten_to_twenty)
print("Now running:$20-50M...")
twenty_to_fifty_total = cmccall.handle(20, 50, twenty_to_fifty)

# Judge long-term
print("Judging long-term holding for $1-10M...")
cmccall.longTerm(one_to_ten, one_to_ten_total)
print("Judging long-term holding for $10-20M...")
cmccall.longTerm(ten_to_twenty, ten_to_twenty_total)
print("Judging long-term holding for $20-50M...")
cmccall.longTerm(twenty_to_fifty, twenty_to_fifty_total)

# Judge underrated
print("Judging if the projects of $1-10M are underrated...")
cmccall.underrated(one_to_ten, one_to_ten_total)
print("Judging if the projects of $10-20M are underrated...")
cmccall.underrated(ten_to_twenty, ten_to_twenty_total)
print("Judging if the projects of $20-50M are underrated...")
cmccall.underrated(twenty_to_fifty, twenty_to_fifty_total)

print("All completed, saving file...")
try:
    route = asksaveasfilename(filetypes=[("Excel file","*.xlsx")])
    if(route == ''):
        raise NotADirectoryError
    wb.save(route)
    print("Done.")
except NotADirectoryError:
    print("Cancelled.")
except:
    print("Something went wrong.")
