import tabula
import numpy as np
import pandas as pd
import openpyxl
#x = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= page_num, relative_area = True,area=(29.4, 0, 49.3 +29.4 ,78))
x = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= 1, relative_area = True,area=(23, 0, 49.3 + 23 ,0 + 110))

#x = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= page_num, relative_area = True,area=(23, 0, 29.4,0 + 110))
#x = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= page_num, relative_area = True)
#y = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages = 1, relative_area = True,area=(23, 0, 49.3 + 23 ,10))

#y = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages = 1, relative_area = True,area=(4.32 , 48, 14.88 , 88.8))  


# Поиск места     
y = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages = 1, relative_area = True,area=(3, 49, 11, 91))        
#CODe
#y = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages = 1, relative_area = True,area=(15, 0, 20, 8))        
#y = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages = 1, relative_area = True,area=(15, 8, 20, 23))
#y = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages = 1, relative_area = True,area=(15, 23, 20, 30))
#y = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages = 1, relative_area = True,area=(15, 23, 20, 45))
#y = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages = 1, relative_area = True,area=(15, 40, 20, 45))

#big tab 

#y1 = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= 1, relative_area = True,area=(21, 0, 72, 11))
#y2 = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= 1, relative_area = True,area=(21, 11, 72, 41))
#SIZE = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= 1, relative_area = True,area=(21, 39, 72, 47))
#UM = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= 1, relative_area = True,area=(21, 46, 72, 52))
#QTY = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= 1, relative_area = True,area=(21, 50, 72, 57))
#UN_PRICE = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= 1, relative_area = True,area=(21, 60, 72, 74))
#DISC = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= 1, relative_area = True,area=(21, 74, 69, 78))
#VALUE = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= 1, relative_area = True,area=(21, 78, 69, 89))
#CI = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= 1, relative_area = True,area=(21, 88, 69, 94))
check = np.array(y)
print(check)

#print(check)

send = np.array(y)

df = pd.DataFrame(send[0])

df.to_excel('example.xlsx', index=False)