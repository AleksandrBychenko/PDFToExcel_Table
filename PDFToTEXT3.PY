import PyPDF2
import textract
import numpy as np
import openpyxl
import os
import tabula

#Проход по страницам
with open("54564-54566.pdf", 'rb') as file:
    path = "file2.xlsx"

    if os.path.exists(path):
        # Открываем файл Excel
        workbook = openpyxl.load_workbook(path)
    else:
        # Создаем новый файл Excel
        workbook = openpyxl.Workbook()
        workbook.save(path)
    
        # Получаем активный лист
    worksheet = workbook.active

    for page_num in range(1):

        page_num = 1 
        send = []   

        send1 = ['CUSTOMER', 'Cust Code', 'VAT NO', 'NO', 'Date', 'Page', 'ITEM Code', 'DESCRIPTION','SIZE', 'U.M', 'QTY', 'UNIT PRICE','VALUE','', 'C.I']

        # Создаем массив данных для новой строки
        #new_row = [6, 7, 8, 9, 10]


        # Вычисляем номер последней строки в таблице
        last_row = worksheet.max_row + 1


        # Добавляем новую строку в таблицу
        for i in range(len(send1)):
            from openpyxl.styles import Font
            bold_font = Font(bold=True)
            worksheet.cell(row= 1, column=i+1, value= send1[i])
            worksheet.cell(row=1, column= i+1).font = bold_font

        
        CUSTOMER = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 49, 11, 91))        
        #CUSTOMER = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 4, 5, 6))        
        
        #print(CUSTOMER)
        send_CUST = np.array(CUSTOMER)
        #print(send_CUST)
        '''
        if send_CUST == []:
            print("Массив не пустой")
        else:
            print("Массив пустой")
        '''
        # создание пустой строки
        my_str= ''

        # проход по элементам массива и добавление каждого элемента в строку
        '''
        for element in send_CUST[0]:
            my_string += element
        print(my_string)
        '''
        for i in range(len(send_CUST)):
            for j in range(len(send_CUST[i])):
                for k in range(len(send_CUST[i][j])):
                    my_str += str(send_CUST[i][j][k]) # добавляем каждый элемент в конец строки
        #stringg  =  CUSTOMER[0]
        #print(stringg)
        # объединение всех элементов массива в одну строку с разделителем ", "
        #joined_array = ' '.join(send_CUST)
        worksheet['A2'] = my_str
        #worksheet.cell(row= 2, column= 1, value= CUSTOMER[0])
       
#<----------------------------------------------------

        all = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= page_num, relative_area = True,area=(21, 0, 49.3 + 23 ,0 + 110))
        send_all = np.array(all)
        #print (send_all)
        sole =''
        row_index = 2
        sole_mas = []
        # Проходим по каждому элементу массива
        for i in range(len(send_all)):
            for j in range(len(send_all[i])):
                for k in range(len(send_all[i][j])):
            # Проверяем наличие слова SOLE в элементе
                        if 'SOLE:' in str(send_all[i][j][k]):
                # Добавляем слово в строку
                            worksheet['D' + str(row_index)] = send_all[i][j][k]
                            # увеличение индекса строки на 1
                            row_index += 1
                            sole += str(send_all[i][j][k]) # добавляем каждый элемент в конец строки
                            new_array = send_all[i][j][k].split()
                            sole_mas.append(new_array)
        print(sole_mas)

#<----------

        y1 = tabula.read_pdf('54564-54566.pdf', stream = True, multiple_tables = False, pages= 1, relative_area = True,area=(22, 0, 72, 11))
        send_y1 = np.array(y1)

        row_index = 2

        nuzno = True
                # Проходим по каждому элементу массива
        for i in range(len(send_y1)):
            for j in range(len(send_y1[i])):
                for k in range(len(send_y1[i][j])):
            # Проверяем наличие слова SOLE в элементе
                    
                    for x in range(len(sole_mas)):
                        for y in range(len(sole_mas[x])):
                            send_delet = str(send_y1[i][j][k]).split(' ')
                            for z in range(len(send_delet)):
                                if sole_mas[x][y] == send_delet[z]:
                                    nuzno = False
                                    print(send_y1[i][j][k])
                            #if not any(word in send_y1[i][j][k] for word in sole.split(' ')):
                            # Добавляем слово в строку
                    
                    if  nuzno  == True:
                        #print(send_y1[i][j][k])
                         #worksheet.cell(row= r + k , column = 3, value = send_y1[i][j][k])
                        worksheet['G' + str(row_index)] = send_y1[i][j][k]
                         # увеличение индекса строки на 1
                        row_index += 1
                         
                         
                         #print('A' + str(row_index))
                         
                         
                         #sole += str(send_all[i][j][k]) # добавляем каждый элемент в конец строки
                    nuzno  = True
        # Сохраняем файл
        
        workbook.save(path)