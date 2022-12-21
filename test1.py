# write data openpyxl refer 
def save_data():
    wb = openpyxl.Workbook()
    ws = wb.active
    print(entryArray)
    for b in range(len(entryArray)):
        temp = entryArray[b]
        ws.cell(row=b+1, column=1).value = temp
    ws.save('/home/'+getpass.getuser()+'/Desktop/FileName.xlsx')


# otherwhie solutiion 
wb = openpyxl.load_workbook("./test_data.xlsx")

wb.create_sheet('data')
sh = wb['data']

sh.cell(1, 1).value = "qavbox"
print(sh.cell(1, 1).value)

wb.save("./test_data.xlsx")
sh['A1'].value = "QAVBOX"
print(sh.cell(1, 1).value)
testdata = [('username', 'password', 'type'), ('abc', 123, 'valid'), ('def', 456, 'invalid')]

for item in testdata:
    sh.append(item)

wb.save("./test_data.xlsx")
testdata1 = ("qavbox", "qavalidation.com")
sh.append(testdata1)