import pandas as pd
import xlwings as xw
import os


for i in range(48):
    if i < 9:
        machine = 'N0{}'.format(i+1)
    else:
        machine = 'N{}'.format(i+1)

    wb = xw.Book('testInput.xlsm')
    ws = wb.sheets['Sheet']
    ws['C2'].value = machine
    wb.save('temp.xlsm')
    wb.close()

    data = pd.read_excel('temp.xlsm', 'ImportKomplet', header = None)

    toPrint = data[0]

    impfile = open('Import/{}'.format(machine), 'w')

    for row in toPrint:
        if type(row) == float:
            impfile.write('\n')
        else:
            impfile.write('{} \n'.format(row))
    impfile.close()

os.remove('test.xlsm')