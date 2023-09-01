import tabula
import numpy as np
import pandas as pd
import openpyxl
#x = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= page_num, relative_area = True,area=(29.4, 0, 49.3 +29.4 ,78))
#x = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= 1, relative_area = True,area=(23, 0, 49.3 + 23 ,0 + 110))

#x = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= page_num, relative_area = True,area=(23, 0, 29.4,0 + 110))
#x = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= page_num, relative_area = True)
#y = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages = 1, relative_area = True,area=(23, 0, 49.3 + 23 ,10))

#y = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages = 1, relative_area = True,area=(4.32 , 48, 14.88 , 88.8))  


# Поиск места     
#y = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages = 1, relative_area = True,area=(3, 49, 11, 91))        
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
y = tabula.read_pdf("54564-54566.pdf", stream = True, multiple_tables = False, pages= 8, relative_area = True,area=(21, 45, 69, 53))
           
check = np.array(y)
print(check)

#print(check)

send = np.array(y)

df = pd.DataFrame(send[0])

df.to_excel('example.xlsx', index=False)

--
                                 # обработка этой строки на нужные слова !!!!!  - не учитывапется если нет sole
                                my_list = send_y1[i][j][k].split('Country of origin:')
                                if len(my_list) == 1: 
                                    my_list = my_list[0].split('H. S')
                                if len(my_list) > 1: 
                                    my_list = my_list[1].split('H. S')
                                
                                if len(my_list) == 2:
                                    worksheet['J' + str(row_index)] = my_list[0].replace(" ", "")
                                    #oтделяем по :
                                    hs = str(my_list[1].replace(":", "")) 
                                    worksheet['K' + str(row_index)] = hs.replace(" ", "")
                                    code = my_list[1].replace(" ", "")
                                    code = code.replace(":", "")
                                if len(my_list) > 2:
                                    #!!!!!!!!!!!!!!!!
                                    hs = str(my_list[0].replace(":", ""))
                                    worksheet['J' + str(row_index)] = hs.replace(" ", "")
                                    code = my_list[len(my_list) - 1].replace(" ", "")
                                    code = code.replace(":", "")
                                
                                if len(my_list) < 2:
                                    #!!!!!!!!!!!!!!!!
                                    hs = str(my_list[0].replace(":", ""))
                                    worksheet['J' + str(row_index)] = hs.replace(" ", "")
                                    code = my_list[0].replace(" ", "")
                                    code = code.replace(":", "")
                                
                                #print(code)
                                '''
                                if cdvig == last_row:
                                    description = send_y1[i][j-2][k]
                                    description_send = description.split(code)
                                    
                                    worksheet['H' + str(row_index)] = description_send[0]
                                else:
                                    
                                    if row_index != last_row:
                                        description = send_y1[i][j-2][k]
                                        description_send = description.split(code)
                                        worksheet['H' + str(row_index)] = description_send[0]
                                '''
                                description = send_y1[i][j-2][k]
                                description_send = description.split(code)
                                worksheet['H' + str(row_index)] = description_send[0]

                                if first_ras and cdvig == last_row + 1:    
                                    description = preded
                                    description_send = description.split(code)
                                    worksheet['H' + str(row_index)] = description_send[0] 
                                    first_ras = False

                                preded  =  send_y1[-1][-1][-1] 