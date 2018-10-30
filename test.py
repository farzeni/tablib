from tablib import Dataset
from datetime import date

d = Dataset()

d.append([date.today(), 'antani'])
d.append([date.today(), 'antani2'])
d.append([date.today(), 'antani3'])

d.headers = ['test', 'stringa']

with open('test.xls', 'wb') as f:
     f.write(d.export('xlsx'))
