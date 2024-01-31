import pandas as pd

data = pd.read_excel("import_gen_v2.xlsm", "ImportKomplet")

toPrint = data.test1.array
impfile = open("test.imp", "w")

for row in toPrint:
    impfile.write("{} \n".format(row))

impfile.close()