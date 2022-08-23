import openpyxl, CMCcall
from openpyxl.styles import Font
from tkinter.filedialog import asksaveasfilename

wb = openpyxl.load_workbook('templatexlsx.xlsx')
yesFont = Font(color='00FF00', bold='True')
noFont = Font(color='FF0000')

one_to_ten = wb.worksheets[0]
ten_to_twenty = wb.worksheets[1]
twenty_to_fifty = wb.worksheets[2]
#$1-10M
print("Now running: $1-10M...")
i = 2
for name in CMCcall.name(1, 10):
    one_to_ten[f'A{i}'] = name
    i += 1
i = 2
for tvl in CMCcall.tvl(1, 10):
    one_to_ten[f'B{i}'] = tvl
    i += 1
i = 2
for change in CMCcall.change(1, 10):
    one_to_ten[f'C{i}'] = change
    i += 1
i = 2
for fdv in CMCcall.fdv(1, 10):
    one_to_ten[f'D{i}'] = fdv
    i += 1
i = 2
for mc in CMCcall.mc(1, 10):
    one_to_ten[f'E{i}'] = mc
    i += 1

one_to_ten_totalnum = i - 2
#$10-20M
print("Now running:$10-20M...")
i = 2
for name in CMCcall.name(10, 20):
    ten_to_twenty[f'A{i}'] = name
    i += 1
i = 2
for tvl in CMCcall.tvl(10, 20):
    ten_to_twenty[f'B{i}'] = tvl
    i += 1
i = 2
for change in CMCcall.change(10, 20):
    ten_to_twenty[f'C{i}'] = change
    i += 1
i = 2
for fdv in CMCcall.fdv(10, 20):
    ten_to_twenty[f'D{i}'] = fdv
    i += 1
i = 2
for mc in CMCcall.mc(10, 20):
    ten_to_twenty[f'E{i}'] = mc
    i += 1
ten_to_twenty_totalnum = i - 2
#$20-50M
print("Now running:$20-50M...")
i = 2
for name in CMCcall.name(20, 50):
    twenty_to_fifty[f'A{i}'] = name
    i += 1
i = 2
for tvl in CMCcall.tvl(20, 50):
    twenty_to_fifty[f'B{i}'] = tvl
    i += 1
i = 2
for change in CMCcall.change(20, 50):
    twenty_to_fifty[f'C{i}'] = change
    i += 1
i = 2
for fdv in CMCcall.fdv(20, 50):
    twenty_to_fifty[f'D{i}'] = fdv
    i += 1
i = 2
for mc in CMCcall.mc(20, 50):
    twenty_to_fifty[f'E{i}'] = mc
    i += 1
twenty_to_fifty_totalnum = i - 2
#Long-term holding
print("Judging long-term holding for $1-10M...")
for t in range(2, one_to_ten_totalnum + 2):
    one_to_ten[f'F{t}'] = (one_to_ten[f'E{t}'].value) * 100 / (one_to_ten[f'D{t}'].value)
    if one_to_ten[f'F{t}'].value > 40:
        one_to_ten[f'F{t}'].font = yesFont
    else:
        one_to_ten[f'F{t}'].font = noFont
        
print("Judging long-term holding for $10-20M...")
for t in range(2, ten_to_twenty_totalnum + 2):
    ten_to_twenty[f'F{t}'] = (ten_to_twenty[f'E{t}'].value) * 100 / (ten_to_twenty[f'D{t}'].value)
    if ten_to_twenty[f'F{t}'].value > 40:
        ten_to_twenty[f'F{t}'].font = yesFont
    else:
        ten_to_twenty[f'F{t}'].font = noFont

print("Judging long-term holding for $20-50M...")
for t in range(2, twenty_to_fifty_totalnum + 2):
    twenty_to_fifty[f'F{t}'] = (twenty_to_fifty[f'E{t}'].value) * 100 / (twenty_to_fifty[f'D{t}'].value)
    if twenty_to_fifty[f'F{t}'].value > 40:
        twenty_to_fifty[f'F{t}'].font = yesFont
    else:
        twenty_to_fifty[f'F{t}'].font = noFont
#Underrated
print("Judging if the projects of $1-10M are underrated...")
for t in range(2, one_to_ten_totalnum + 2):
    one_to_ten[f'G{t}'] = (one_to_ten[f'E{t}'].value) / (one_to_ten[f'B{t}'].value)
    if one_to_ten[f'G{t}'].value < 1:
        one_to_ten[f'G{t}'].font = yesFont
    else:
        one_to_ten[f'G{t}'].font = noFont
print("Judging if the projects of $10-20M are underrated...")
for t in range(2, ten_to_twenty_totalnum + 2):
    ten_to_twenty[f'G{t}'] = (ten_to_twenty[f'E{t}'].value) / (ten_to_twenty[f'B{t}'].value)
    if ten_to_twenty[f'G{t}'].value < 1:
        ten_to_twenty[f'G{t}'].font = yesFont
    else:
        ten_to_twenty[f'G{t}'].font = noFont
print("Judging if the projects of $20-50M are underrated...")
for t in range(2, twenty_to_fifty_totalnum + 2):
    twenty_to_fifty[f'G{t}'] = (twenty_to_fifty[f'E{t}'].value) / (twenty_to_fifty[f'B{t}'].value)
    if twenty_to_fifty[f'G{t}'].value < 1:
        twenty_to_fifty[f'G{t}'].font = yesFont
    else:
        twenty_to_fifty[f'G{t}'].font = noFont
print("All completed, saving file...")
try:
    route = asksaveasfilename(filetypes=[("Excel file","*.xlsx")])
    wb.save(route + ".xlsx")
    print("Done.")
except:
    print("Cancelled.")
