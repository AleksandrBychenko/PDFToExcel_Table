import PyPDF2
import textract
import numpy as np
import openpyxl
import math
import os
import tabula
import pandas as pd

#Проход по страницам
#<-------------------

#file = "54564-54566.pdf"
#file = "Example.pdf"
file = "2.pdf"


def fromPDftoExelspec(file):
    # Вычисляем номер последней строки в таблице
    last_row = 2

    #весь столбик 
    full_um = []

    len_size = 0

    # начальное значение предыдущегог числа на рандом выставил 
    preded = str(2)
# открываем пдф и проходим по страницам 
    with open(file, 'rb') as file:
        
        pdf_reader = PyPDF2.PdfReader(file)

        page_number = len(pdf_reader.pages)
        
        #начаьные страницы ?????????????
                     #сканируем область      
        nO = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = 1, relative_area = True,area=(15, 23, 20, 30))        
        
        #проверка если место будет пустое 
        #Cust_Code = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 4, 5, 6))        
        
        #делаем из отосланного массив
        send_CUST = np.array(nO)
    
        # создание пустой строки
        my_str= ''
    
        # проход по элементам массива и добавление каждого элемента в строку
        for i in range(len(send_CUST)):
            for j in range(len(send_CUST[i])):
                for k in range(len(send_CUST[i][j])):
                    my_str += str(send_CUST[i][j][k]) # добавляем каждый элемент в конец строки
    
        
        ####### end check ::::::::

        
        #path = "file2.xlsx"
        path_you  = 'scan_'+str(my_str)+'.xlsx'
        #???????????????

        #for page_num in range(6,8):
        for page_num in range(1,(len(pdf_reader.pages) + 1)):
            #page_num = 3

            #####Проверка на номер строки  ::::::::::::::::::::::::
                        #сканируем область      
             #сканируем область      
            nO = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(15, 23, 20, 30))        
            
            #проверка если место будет пустое 
            #Cust_Code = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 4, 5, 6))        
            
            #делаем из отосланного массив
            send_CUST = np.array(nO)
    
            # создание пустой строки
            my_str= ''
    
            # проход по элементам массива и добавление каждого элемента в строку
            for i in range(len(send_CUST)):
                for j in range(len(send_CUST[i])):
                    for k in range(len(send_CUST[i][j])):
                        my_str += str(send_CUST[i][j][k]) # добавляем каждый элемент в конец строки
    
            
            ####### end check ::::::::

            
            #path = "file2.xlsx"
            path = 'scan_'+str(my_str)+'.xlsx'

            if  path_you != path:
                last_row = 2
                path_you = path
                #весь столбик 
                full_um = []
                len_size = 0
            else:
                pass
    
            if os.path.exists(path):
                # Открываем файл Excel
                workbook = openpyxl.load_workbook(path)
            else:
                # Создаем новый файл Excel
                workbook = openpyxl.Workbook()
                #workbook.save(path)
             
                 # Получаем активный лист
            worksheet = workbook.active
    
             # Очищаем содержимое листа
            #worksheet.delete_rows(1, worksheet.max_row)
    
            #сразуже номер для проверки
            #page_num = 11
            send = []   
            # Вычисляем номер последней строки в таблице
            #last_row = worksheet.max_row + 1
    #<-------------------
    
    #<1@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    # название столбцов 
            send1 = ['CUSTOMER', 'Cust Code', 'VAT NO', 'NO', 'Date', 'Page', 'ITEM Code', 'DESCRIPTION', 'COMPOSITION','COUNTRY OF ORIGIN',	'CUSTOMS CODE (H.S.)','SIZE', 'U.M', 'QTY', 'UNIT PRICE','VALUE', 'C.I', '% DISCOUNT',	'GROSS VALUE',	'NET VALUE',	'ADVANCE', 	'TOTAL VALUE',	'TOTAL INVOICE',	'TOTAL AMOUNT DUE',	'PACKAGING NUMBER',	'GROSS WEIGHT', 'NET WEIGHT', 'VOLUME']
    

    
            # Добавляем название столбцов 
            for i in range(len(send1)):
                from openpyxl.styles import Font
                bold_font = Font(bold=True)
                worksheet.cell(row= 1, column=i+1, value= send1[i])
                worksheet.cell(row=1, column= i+1).font = bold_font
    #<1концец---------------------------------
    
    
    
    #<2Castomeer---------------------------------
             #сканируем область       
            CUSTOMER = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 49, 11, 91))        
            
            #проверка если место будет пустое 
            #CUSTOMER = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 4, 5, 6))        
            
            #делаем из отосланного массив
            send_CUST = np.array(CUSTOMER)
    
            # создание пустой строки
            my_str= ''
    
            # проход по элементам массива и добавление каждого элемента в строку
            for i in range(len(send_CUST)):
                for j in range(len(send_CUST[i])):
                    for k in range(len(send_CUST[i][j])):
                        my_str += str(send_CUST[i][j][k]) # добавляем каждый элемент в конец строки
    
            #добовляем значение
            #row_numb = str(last_row)
            row_numb = 2
            worksheet['A'+ str(row_numb)] = my_str
     
    #<2Castomeer-end--------------------------------
    
    #<2Cust Code---------------------------------
            #сканируем область      
            Cust_Code = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(15, 0, 20, 8))        
            
            #проверка если место будет пустое 
            #Cust_Code = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 4, 5, 6))        
            
            #делаем из отосланного массив
            send_CUST = np.array(Cust_Code)
    
            # создание пустой строки
            my_str= ''
    
            # проход по элементам массива и добавление каждого элемента в строку
            for i in range(len(send_CUST)):
                for j in range(len(send_CUST[i])):
                    for k in range(len(send_CUST[i][j])):
                        my_str += str(send_CUST[i][j][k]) # добавляем каждый элемент в конец строки
    
            #добовляем значение
            #row_numb = str(last_row)
            row_numb = 2
            worksheet['B' + str(row_numb)] = my_str
     
    #<Cust Code-end--------------------------------
    
    #<VAT NO---------------------------------
            #сканируем область      
            VAT_nO = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(15, 8, 20, 23))        
            
            #проверка если место будет пустое 
            #Cust_Code = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 4, 5, 6))        
            
            #делаем из отосланного массив
            send_CUST = np.array(VAT_nO)
    
            # создание пустой строки
            my_str= ''
    
            # проход по элементам массива и добавление каждого элемента в строку
            for i in range(len(send_CUST)):
                for j in range(len(send_CUST[i])):
                    for k in range(len(send_CUST[i][j])):
                        my_str += str(send_CUST[i][j][k]) # добавляем каждый элемент в конец строки
    
            #добовляем значение
            #row_numb = str(last_row)
            row_numb = 2
            worksheet['C' +str(row_numb)] = my_str
     
    #<VAT NO-end--------------------------------
    
    
    #<NO---------------------------------
            #сканируем область      
            nO = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(15, 23, 20, 30))        
            
            #проверка если место будет пустое 
            #Cust_Code = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 4, 5, 6))        
            
            #делаем из отосланного массив
            send_CUST = np.array(nO)
    
            # создание пустой строки
            my_str= ''
    
            # проход по элементам массива и добавление каждого элемента в строку
            for i in range(len(send_CUST)):
                for j in range(len(send_CUST[i])):
                    for k in range(len(send_CUST[i][j])):
                        my_str += str(send_CUST[i][j][k]) # добавляем каждый элемент в конец строки
    
            #добовляем значение
            #row_numb = str(last_row)
            row_numb = 2
            worksheet['D' + str(row_numb)] = my_str
            No = my_str
     
    #<NO-end--------------------------------
    
    #<Date--Page-------------------------------
    
    
            #сканируем область    !!! в данной облости можно только 3 ячейки сразу же !!!!!!!!!!!!!!!!!!!  
            date = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(15, 23, 20, 47))        
            
            #проверка если место будет пустое 
            #date = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 4, 5, 6))        
            
            #делаем из отосланного массив
            send_CUST = np.array(date)

            # обозначаем номер строчки 
            #row_numb = str(last_row)
            row_numb = 2
    
            #просматриваем трайной массив и отправляем его по частям 
            if not date:
                pass
            else:
                if len(send_CUST[0][0]) == 3:
                    worksheet['E' + str(row_numb)] = send_CUST[0][0][1]
                    worksheet['F' + str(row_numb)] = send_CUST[0][0][2]
                if len(send_CUST[0][0]) == 2 and send_CUST[0][0][0] == '54564':
                    worksheet['E' + str(row_numb)] = send_CUST[0][0][1]
                    worksheet['F' + str(row_numb)] = send_CUST[0][0][2]
                if len(send_CUST[0][0]) == 1 and send_CUST[0][0][0] == '54564':
                    worksheet['E' + str(row_numb)] = send_CUST[0][0][1]
                    
            
    #<Date-Page-end--------------------------------
    
    #<3SOLE:--!!!! и дрогое--------------------------------------------------
            #1)Проходимся по всему участку 2) делаем массив исклучений sole_mas 3) добавляем Sole  
            all = tabula.read_pdf(file, stream = True, multiple_tables = False, pages= page_num, relative_area = True,area=(21, 0, 69,0 + 110))
            send_all = np.array(all)
            sole =''
            row_index = last_row
            sole_mas = []

            #проверка на % в строке     
            y1 = tabula.read_pdf(file, stream = True, multiple_tables = False, pages= page_num, relative_area = True,area=(22, 0, 69, 11))
            send_y1 = np.array(y1)
            prov_sol = []
                    # Проходим по каждому элементу массива
            for i in range(len(send_y1)):
                for j in range(len(send_y1[i])):
                    for k in range(len(send_y1[i][j])):
                        prov_sol.append(send_y1[i][j][k])
            ostan = False
            # Проходим по каждому элементу массива
            for i in range(len(send_all)):
                for j in range(len(send_all[i])):
                    for k in range(len(send_all[i][j])):
                # Проверяем наличие слова SOLE в элементе
                            '''
                            if 'SOLE:' in str(send_all[i][j][k]):
                                # Добавляем слово в строку
                                worksheet['I' + str(row_index)] = send_all[i][j][k]
                                # увеличение индекса строки на 1
                                row_index += 1
                                sole += str(send_all[i][j][k]) # добавляем каждый элемент в конец строки
                                new_array = send_all[i][j][k].split()
                                sole_mas.append(new_array)
                            '''
                            for w in range(len(prov_sol)):
                                if prov_sol[w] in str(send_all[i][j][k]) and '%' in str(send_all[i][j][k]) and ostan == False:
                                    # Добавляем слово в строку
                                    worksheet['I' + str(row_index)] = send_all[i][j][k]
                                    # увеличение индекса строки на 1
                                    
                                    sole += str(send_all[i][j][k]) # добавляем каждый элемент в конец строки
                                    new_array = send_all[i][j][k].split()
                                    sole_mas.append(new_array)
                                    row_index += 1
                                    ostan = True
                            ostan = False
            #print(sole_mas)
    #<3SOLE:-end-----COMPOSITION----------------------------------------------
    
    #провверка не другой ли формат $$$$$
            net_li = False
    
    #<проверка на сдвиг-------- 
            #проверка сдвиг есть или нет относительно начала ITEM CODE
            cdvig = last_row + 1
            #сканируем область 
            y1 = tabula.read_pdf(file, stream = True, multiple_tables = False, pages= page_num, relative_area = True,area=(22, 0, 69, 11))
            #записываем в массив 
            #y1 = tabula.read_pdf(file, stream = True, multiple_tables = False, pages= page_num, relative_area = True,area=(5, 0, 10, 12))
            
            send_y1 = np.array(y1)
            
            #провверка не другой ли формат $$$$$
            if not y1:
                net_li = True
    
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
                                        
                                        if j == 0:
                                            #3
                                            cdvig = last_row + 1
                                        elif j == 1:
                                            cdvig = last_row 
    
    
    #<проверка на сдвиг--end------ 
    
    
    #<4-ITEM CODE----------
            #>>>это поменять y1 !!!!!!!!!
            #сканируем область 
            y1 = tabula.read_pdf(file, stream = True, multiple_tables = False, pages= page_num, relative_area = True,area=(22, 0, 69, 11))
            #записываем в массив 
            #>>>это поменять y1 !!!!!!!!!
            send_y1 = np.array(y1)
            
            #проверка на сдвиг
     
            row_index = cdvig
    
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
                                        
                                    
                                    #проверка не совпадает ли с массивом проверки 
                                    if sole_mas[x][y] == send_delet[z]:
                                        nuzno = False
                                    #проверка на пустой     
                                    if send_y1[i][j][k] is  None:
                                        nuzno = False
                                    
                        if  nuzno  == True:
    
                            worksheet['G' + str(row_index)] = send_y1[i][j][k]
                             # увеличение индекса строки на 1
                            #len_size += 1
                            row_index += 1
                            len_size = row_index
    
                        nuzno  = True
            
    #<4-ITEM CODE-end---------
    
    #<DESCRIPTION-COMPOSITIO---COUNTRY OF ORIGIN---------
            #сканируем участок 
            y2 = tabula.read_pdf(file, stream = True, multiple_tables = False, pages= page_num, relative_area = True,area=(22, 11, 69, 41))
            #то что отсканируем переводим в массив
            send_y1 = np.array(y2)
            
            
            #c какой стороки начинаем 
            row_index = last_row 
        

            #проверка что есть элементы
            pusto = False
    
            #разделитель для колонки описания 
            code =' '
    
            #проверка на первый раз  использования
            first_ras = True
    
            #проверка на то что мы отсканировали нужно ли вставлять 
            nuzno = True
            # Проходим по каждому элементу массива отсканированному 
            for i in range(len(send_y1)):
                for j in range(len(send_y1[i])):
                    for k in range(len(send_y1[i][j])):
               
                        #проходим по массиву проверки и сравниваем 
                        for x in range(len(sole_mas)):
                            for y in range(len(sole_mas[x])):
    
    
                                send_delet = str(send_y1[i][j][k]).split(' ')
                                #проходимся по строке элемента сканера                             
                                for z in range(len(send_delet)):
    
                                    #проверка не совпадает ли с массивом проверки 
                                    if sole_mas[x][y] == send_delet[z]:
                                        nuzno = False
                                        
                                    #проверка на пустой     
                                    if send_y1[i][j][k] is  None:
                                        nuzno = False
                                    
                                    #проверка на нан , часто пападается (это работает если столбец не пустой проверки!)
                                    if str(send_y1[i][j][k]) == 'nan':
                                        nuzno = False
                                        
                              
    
                        #если подходит отсылаем 
                        if  nuzno  == True:
                            #находим строчку со страной и от нее и скачим 
                            if   'Country' in str(send_y1[i][j][k]):
                                
                                my_list = send_y1[i][j][k].split('H. S')
                                #print(my_list)
                                if len(my_list) > 1:
                                    hs  = str(my_list[1].replace(":", "")) 
                                    worksheet['K' + str(row_index)] = hs.replace(" ", "")
                                    code = hs.replace(" ", "")
                                else:
                                    my_list = my_list[0].split('H.S')
                                    if len(my_list) > 1:
                                        hs  = str(my_list[1].replace(":", "")) 
                                        worksheet['K' + str(row_index)] = hs.replace(" ", "")
                                        code = hs.replace(" ", "")

                                start = send_y1[i][j][k].find('Country of origin:') + len('Country of origin:')
                                end = send_y1[i][j][k].find('H.', start)

                                if start != -1 and end != -1:
                                    country = send_y1[i][j][k][start:end].strip()
                                    cn = str(country.replace(" ", ""))
                                    worksheet['J' + str(row_index)] = cn
                                
                                if first_ras and cdvig == last_row + 1:    
                                    description = preded
                                    description_send = description.split(code)
                                    worksheet['H' + str(row_index)] = description_send[0] 
                                    first_ras = False
                                else:
                                    description = send_y1[i][j-2][k]
                                    description_send = description.split(code)
                                    worksheet['H' + str(row_index)] = description_send[0] 
                                    first_ras = False


                                # увеличение индекса строки на 1 
                                row_index += 1
                                
                            '''    
                            if 'Deposit:' in str(send_y1[i][j][k]):
                                    net_li = True
                            '''
                        # возрощаем значение для проверки
                        nuzno  = True
                         
                        

                        '''
                        #провверка не другой ли формат $$$$$
                        if net_li == True:
                             #проверка на нан , часто пападается (это работает если столбец не пустой проверки!)
                            if str(send_y1[i][j][k]) == 'nan':
                                pass
                            else:
                                worksheet['H' + str(row_index)] = send_y1[i][j][k]
                                row_index += 1
                        '''        
            preded  =  send_y1[-1][-1][-1] 
            net_li = False              
    
    #<DESCRIPTION-COMPOSITIO---COUNTRY OF ORIGIN----end-----
    
    #<-5--SIZE--(расписано)------------
            #сканируем участок 
            SIZE = tabula.read_pdf(file, stream = True, multiple_tables = False, pages= page_num, relative_area = True,area=(21, 39, 69, 47))
            #то что отсканируем переводим в массив
            send_y1 = np.array(SIZE)
            
            #c какой стороки начинаем 
            row_index = cdvig

            
            #проверка на то что мы отсканировали нужно ли вставлять 
            nuzno = True
            # Проходим по каждому элементу массива отсканированному 
            for i in range(len(send_y1)):
                for j in range(len(send_y1[i])):
                    for k in range(len(send_y1[i][j])):
               
                        #проходим по массиву проверки и сравниваем 
                        for x in range(len(sole_mas)):
                            for y in range(len(sole_mas[x])):
    
    
                                send_delet = str(send_y1[i][j][k]).split(' ')
                                #проходимся по строке элемента сканера                             
                                for z in range(len(send_delet)):
    
                                    #проверка не совпадает ли с массивом проверки 
                                    if sole_mas[x][y] == send_delet[z]:
                                        nuzno = False
                                    #проверка на пустой     
                                    if send_y1[i][j][k] is  None:
                                        nuzno = False
                                    
                                    #проверка на нан , часто пападается (это работает если столбец не пустой проверки!)
                                    if str(send_y1[i][j][k]) == 'nan':
                                        nuzno = False
                                    
                                    #проверка на содеражание буквы 
                                    buka = str(send_y1[i][j][k])
                                    if any(c.isalpha() for c in buka):
                                        nuzno = False
                              
    
                        #если подходит отсылаем 
                        if  nuzno  == True:
                            worksheet['L' + str(row_index)] = send_y1[i][j][k]
                             # увеличение индекса строки на 1
                            #len_size += 1
                            row_index += 1
                        #обновляем значение
                        nuzno  = True

    #<-5--SIZE-end-------------        
            
    #<-U.M--------------
            #проверка нет ли sole 
            y1 = tabula.read_pdf(file, stream = True, multiple_tables = False, pages= page_num, relative_area = True,area=(22, 0, 69, 11))
            send_y1 = np.array(y1)
            
            #провверка не другой ли формат $$$$$
            if not y1:
                stvig = 1
            else:
                stvig = 2

            #сканируем участок 
            um = tabula.read_pdf(file, stream = True, multiple_tables = False, pages= page_num, relative_area = True,area=(21, 45, 69, 53))
            #то что отсканируем переводим в массив
            send_y1 = np.array(um)
            #print(send_y1)
            #print(sole_mas)
            #c какой стороки начинаем 
            row_index = cdvig


            #то что главное 
            in_um = [] 

            '''
            #проверка на сдвиг
            if cdvig == 2:
                nachalo = 0
            else:
                nachalo = 1
    
    
            #проверка на то что мы отсканировали нужно ли вставлять nachalo, len(send_y1[i]), stvig 
            nuzno = True
            # Проходим по каждому элементу массива отсканированному 
            for i in range(len(send_y1)):
                for j in range(nachalo, len(send_y1[i]), stvig):
                
                    for k in range(len(send_y1[i][j])):
    
                        #проверка на пустой     
                        if send_y1[i][j][k] is  None:
                            j += 1
    
                        #проверка на нан , часто пападается (это работает если столбец не пустой проверки!)
                        if str(send_y1[i][j][k]) == 'nan':
                            j += 1
                        
     
                        worksheet['M' + str(row_index)] = send_y1[i][j][k]
                         # увеличение индекса строки на 1
                        row_index += 1
                        #обновляем значение
            '''
            #lenn  = len(send_y1[0]) - 1
            #worksheet['M' + str(row_index)] = send_y1[0][lenn][0]            
            #проверка на то что мы отсканировали нужно ли вставлять 
            nuzno = False
            

            # Проходим по каждому элементу массива отсканированному 
            for i in range(len(send_y1)):
                for j in range(len(send_y1[i])):
                    for k in range(len(send_y1[i][j])):
                        
                        if '_' in str(send_y1[i][j][k]):
                            nuzno = True

                        if nuzno  == False:
                            full_um.append(send_y1[i][j][k])
                        


    
    
                        #если подходит отсылаем  worksheet['M' + str(row_index)] = send_y1[i][j][k] row_index += 1
                        if  nuzno  == True:
                            in_um.append(send_y1[i][j][k])
                            # увеличение индекса строки на 1
                        
                        #обновляем значение

            #куда первая запись
            colich = 0
            zapis = 2
            if nuzno == True:
                for i in range(len(full_um)):

                    for j in range(1, len(in_um)):
                        #print(page_num)
                        #print(full_um[i])
                        if i == 0:
                            if  full_um[i] == in_um[j]:
                                colich += 1
                                worksheet['M' + str(zapis)] = full_um[i]
                                zapis = zapis + 1
                        else:
                            if  full_um[i] == in_um[j] and full_um[i] != full_um[i - 1] :
                                colich += 1            
                                worksheet['M' + str(zapis)] = full_um[i]
                                zapis = zapis + 1
                            
                for j in range(len(in_um)):
                    worksheet['M' + str(len_size)] = in_um[j]
                    len_size = len_size + 1
                #print(colich)


    #<-U.M-end-------------          
            
    #<-QTY--------------
            #сканируем участок 
            qty = tabula.read_pdf(file, stream = True, multiple_tables = False, pages= page_num, relative_area = True,area=(21, 50, 72, 61))
            #то что отсканируем переводим в массив
            send_y1 = np.array(qty)
            
            #print(send_y1)
            #c какой стороки начинаем 
            row_index = cdvig
    
            #проверка на то что мы отсканировали нужно ли вставлять 
            nuzno = True
            # Проходим по каждому элементу массива отсканированному 
            for i in range(len(send_y1)):
                for j in range(len(send_y1[i])):
                    for k in range(len(send_y1[i][j])):
               
                        #проходим по массиву проверки и сравниваем 
                        for x in range(len(sole_mas)):
                            for y in range(len(sole_mas[x])):
    
    
                                send_delet = str(send_y1[i][j][k]).split(' ')
                                #проходимся по строке элемента сканера                             
                                for z in range(len(send_delet)):
                                    #print("Мой массив внизу")
                                    #print(send_delet[z])
                                    #print("Проверяющий массив внизу")
                                    #print(sole_mas[x][y])
                                    #проверка не совпадает ли с массивом проверки 
                                    if sole_mas[x][y] == send_delet[z]:
                                        nuzno = False
                                    #проверка на пустой     
                                    if send_y1[i][j][k] is  None:
                                        nuzno = False
                                    
                                    #проверка на нан , часто пападается (это работает если столбец не пустой проверки!)
                                    if str(send_y1[i][j][k]) == 'nan':
                                        nuzno = False
                                     #проверка на содеражание буквы 
                                    buka = str(send_y1[i][j][k])
                                    if any(c.isalpha() for c in buka):
                                        nuzno = False  
                              
    
                        #если подходит отсылаем 
                        if  nuzno  == True:
                            worksheet['N' + str(row_index)] = send_y1[i][j][k]
                             # увеличение индекса строки на 1
                            row_index += 1
                        #обновляем значение
                        nuzno  = True
                               
            
    #<-QTY-end-------------         
            
    #<-UNIT PRICE--------------
            #сканируем участок 
            qty = tabula.read_pdf(file, stream = True, multiple_tables = False, pages= page_num, relative_area = True,area=(21, 60, 71, 72))
            #то что отсканируем переводим в массив
            send_y1 = np.array(qty)
            
            #print(send_y1)
            #c какой стороки начинаем 
            row_index = cdvig
    
            #проверка на то что мы отсканировали нужно ли вставлять 
            nuzno = True
            # Проходим по каждому элементу массива отсканированному 
            for i in range(len(send_y1)):
                for j in range(len(send_y1[i])):
                    for k in range(len(send_y1[i][j])):
               
                        #проходим по массиву проверки и сравниваем 
                        for x in range(len(sole_mas)):
                            for y in range(len(sole_mas[x])):
    
    
                                send_delet = str(send_y1[i][j][k]).split(' ')
                                #проходимся по строке элемента сканера                             
                                for z in range(len(send_delet)):
                                    #print("Мой массив внизу")
                                    #print(send_delet[z])
                                    #print("Проверяющий массив внизу")
                                    #print(sole_mas[x][y])
                                    #проверка не совпадает ли с массивом проверки 
                                    if sole_mas[x][y] == send_delet[z]:
                                        nuzno = False
                                    #проверка на пустой     
                                    if send_y1[i][j][k] is  None:
                                        nuzno = False
                                    
                                    #проверка на нан , часто пападается (это работает если столбец не пустой проверки!)
                                    if str(send_y1[i][j][k]) == 'nan':
                                        nuzno = False
                                    
                                    #проверка на содеражание буквы 
                                    buka = str(send_y1[i][j][k])
                                    if any(c.isalpha() for c in buka):
                                        nuzno = False
                                    
                                        
                              
    
                        #если подходит отсылаем 
                        if  nuzno  == True:
                            worksheet['O' + str(row_index)] = send_y1[i][j][k]
                             # увеличение индекса строки на 1
                            row_index += 1
                        #обновляем значение
                        nuzno  = True
                               
            
    #<-UNIT PRICE-end-------------        
            
    #<-VALUE--------------
            #сканируем участок 
            qty = tabula.read_pdf(file, stream = True, multiple_tables = False, pages= page_num, relative_area = True,area=(21, 77, 69, 89))
            #то что отсканируем переводим в массив
            send_y1 = np.array(qty)
            
            #print(send_y1)
            #c какой стороки начинаем 
            row_index = cdvig
    
            #проверка на то что мы отсканировали нужно ли вставлять 
            nuzno = True
            # Проходим по каждому элементу массива отсканированному 
            for i in range(len(send_y1)):
                for j in range(len(send_y1[i])):
                    for k in range(len(send_y1[i][j])):
               
                        #проходим по массиву проверки и сравниваем 
                        for x in range(len(sole_mas)):
                            for y in range(len(sole_mas[x])):
    
    
                                send_delet = str(send_y1[i][j][k]).split(' ')
                                #проходимся по строке элемента сканера                             
                                for z in range(len(send_delet)):
                                    #print("Мой массив внизу")
                                    #print(send_delet[z])
                                    #print("Проверяющий массив внизу")
                                    #print(sole_mas[x][y])
                                    #проверка не совпадает ли с массивом проверки 
                                    if sole_mas[x][y] == send_delet[z]:
                                        nuzno = False
                                    #проверка на пустой     
                                    if send_y1[i][j][k] is  None:
                                        nuzno = False
                                    
                                    #проверка на нан , часто пападается (это работает если столбец не пустой проверки!)
                                    if str(send_y1[i][j][k]) == 'nan':
                                        nuzno = False
                                    
                                    #проверка на содеражание буквы 
                                    buka = str(send_y1[i][j][k])
                                    if any(c.isalpha() for c in buka):
                                        nuzno = False
                                    
                                        
                              
    
                        #если подходит отсылаем 
                        if  nuzno  == True:
                            worksheet['P' + str(row_index)] = send_y1[i][j][k]
                             # увеличение индекса строки на 1
                            row_index += 1
                        #обновляем значение
                        nuzno  = True
                               
            
    #<-VALUE-end-------------
    
    #<-C.I--------------
            #сканируем участок 
            column_names = ["CI"]
            ci = tabula.read_pdf(file, stream = True, multiple_tables = False, pages= page_num, relative_area = True, area=(21, 87, 69, 94))
            #то что отсканируем переводим в массив
            send_y1 = np.array(ci)
            #проверка на то что мы отсканировали нужно ли вставлять 
            nuzno = True
            
            row_index = cdvig
    
            if not ci:
                pass
            else:
                #ищем заголовок
                df = pd.DataFrame(ci[0])
                
                # получаем название столбца
                column_name = df.columns[0]
    
                #обработка названия столбца 
                if any(c.isalpha() for c in column_name):
                    row_index = cdvig
                else:
                    worksheet['Q' + str(cdvig)] = column_name
                    row_index = cdvig + 1
            
            # Проходим по каждому элементу массива отсканированному 
            for i in range(len(send_y1)):
                for j in range(len(send_y1[i])):
                    for k in range(len(send_y1[i][j])):
               
                        #проходим по массиву проверки и сравниваем 
                        for x in range(len(sole_mas)):
                            for y in range(len(sole_mas[x])):
    
    
                                send_delet = str(send_y1[i][j][k]).split(' ')
                                #проходимся по строке элемента сканера                             
                                for z in range(len(send_delet)):
                                    #print("Мой массив внизу")
                                    #print(send_delet[z])
                                    #print("Проверяющий массив внизу")
                                    #print(sole_mas[x][y])
                                    #проверка не совпадает ли с массивом проверки 
                                    if sole_mas[x][y] == send_delet[z]:
                                        nuzno = False
                                    #проверка на пустой     
                                    if send_y1[i][j][k] is  None:
                                        nuzno = False
                                    
                                    #проверка на нан , часто пападается (это работает если столбец не пустой проверки!)
                                    if str(send_y1[i][j][k]) == 'nan':
                                        nuzno = False
                                    
                                    #проверка на содеражание буквы 
                                    buka = str(send_y1[i][j][k])
                                    if any(c.isalpha() for c in buka):
                                        nuzno = False
                                    
                                        
                              
    
                        #если подходит отсылаем 
                        if  nuzno  == True:
                            worksheet['Q' + str(row_index)] = send_y1[i][j][k]
                             # увеличение индекса строки на 1
                            row_index += 1
                        #обновляем значение
                        nuzno  = True
                               
            
    #<-C.I-end-------------

#2part------------@@@**************************************


#<-GROSS VALUE--------------

             #сканируем область       
            gv = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(67, 0, 70, 15))        
            
            #проверка если место будет пустое 
            #CUSTOMER = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 4, 5, 6))        
            
            #делаем из отосланного массив
            send_CUST = np.array(gv)
    
            my_value = ''
    
            # проход по элементам массива и добавление каждого элемента в строку
            for i in range(len(send_CUST)):
                for j in range(len(send_CUST[i])):
                    for k in range(len(send_CUST[i][j])):
                        #print (send_CUST[i][j][k])
                        my_value = str(send_CUST[i][j][k]) # добавляем каждый элемент в конец строки
    
            #добовляем значение
            worksheet['S2'] = my_value

#<-GROSS VALUE-end-------------

#<-% DISCOUNT--------------

             #сканируем область       
            gv = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(67, 16, 70, 20))        
            
            #проверка если место будет пустое 
            #CUSTOMER = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 4, 5, 6))        
            
            #делаем из отосланного массив
            send_CUST = np.array(gv)
    
            my_value = ''

            # проход по элементам массива и добавление каждого элемента в строку
            for i in range(len(send_CUST)):
                for j in range(len(send_CUST[i])):
                    for k in range(len(send_CUST[i][j])):
                        #print (send_CUST[i][j][k])
                        my_value = str(send_CUST[i][j][k]) # добавляем каждый элемент в конец строки
    
            #добовляем значение
            worksheet['R2'] = my_value

#<-% DISCOUNT-end-------------

#<-% NET VALUE--+Disc Value------------

             #сканируем область    Disc Value   
            #gv = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(67, 21, 70, 35))        
            
            #сканируем область       
            gv = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(67, 36, 70, 49))        
            
            #проверка если место будет пустое 
            #CUSTOMER = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 4, 5, 6))        
            
            #делаем из отосланного массив
            send_CUST = np.array(gv)
    
            my_value = ''

            # проход по элементам массива и добавление каждого элемента в строку
            for i in range(len(send_CUST)):
                for j in range(len(send_CUST[i])):
                    for k in range(len(send_CUST[i][j])):
                        #print (send_CUST[i][j][k])
                        my_value = str(send_CUST[i][j][k]) # добавляем каждый элемент в конец строки
    
            #добовляем значение
            worksheet['T2'] = my_value

#<-% NET VALUE-end-------------

#<-% ADVANCE --------------

                
            
            #сканируем область       
            gv = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(67, 75, 70, 85))        
            
            #проверка если место будет пустое 
            #CUSTOMER = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 4, 5, 6))        
            
            #делаем из отосланного массив
            send_CUST = np.array(gv)
    
            my_value = ''

            # проход по элементам массива и добавление каждого элемента в строку
            for i in range(len(send_CUST)):
                for j in range(len(send_CUST[i])):
                    for k in range(len(send_CUST[i][j])):
                        #print (send_CUST[i][j][k])
                        my_value = str(send_CUST[i][j][k]) # добавляем каждый элемент в конец строки
    
            #добовляем значение
            worksheet['U2'] = my_value

#<-% ADVANCE -end-------------

#<-% PACKAGING NUMBER --------------

                
            
            #сканируем область       
            #gv = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(72, 0, 80, 94))        
            gv = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(81, 0, 85, 7))        
            
            #проверка если место будет пустое 
            #CUSTOMER = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 4, 5, 6))        
            
            #делаем из отосланного массив
            send_CUST = np.array(gv)
    
            my_value = ''

            # проход по элементам массива и добавление каждого элемента в строку
            for i in range(len(send_CUST)):
                for j in range(len(send_CUST[i])):
                    for k in range(len(send_CUST[i][j])):
                        #print (send_CUST[i][j][k])
                        my_value = str(send_CUST[i][j][k]) # добавляем каждый элемент в конец строки
    
            #добовляем значение
            worksheet['Y2'] = my_value

#<-% PACKAGING NUMBER -end-------------

#<-% GROSS WEIGHT --------------

                
            
            #сканируем область       
            #gv = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(72, 0, 80, 94))        
            gv = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(81, 27, 85, 40) )       
            
            #проверка если место будет пустое 
            #CUSTOMER = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 4, 5, 6))        
            
            #делаем из отосланного массив
            send_CUST = np.array(gv)
    
            my_value = ''

            # проход по элементам массива и добавление каждого элемента в строку
            for i in range(len(send_CUST)):
                for j in range(len(send_CUST[i])):
                    for k in range(len(send_CUST[i][j])):
                        #print (send_CUST[i][j][k])
                        my_value = str(send_CUST[i][j][k]) # добавляем каждый элемент в конец строки

            #добовляем значение
            worksheet['Z2'] = my_value

#<-% GROSS WEIGHT -end-------------

#<-% GROSS WEIGHT --------------

                
            
            #сканируем область       
            #gv = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(72, 0, 80, 94))        
            gv = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(81, 27, 85, 40) )       
            
            #проверка если место будет пустое 
            #CUSTOMER = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 4, 5, 6))        
            
            #делаем из отосланного массив
            send_CUST = np.array(gv)
    
            my_value = ''

            # проход по элементам массива и добавление каждого элемента в строку
            for i in range(len(send_CUST)):
                for j in range(len(send_CUST[i])):
                    for k in range(len(send_CUST[i][j])):
                        #print (send_CUST[i][j][k])
                        my_value = str(send_CUST[i][j][k]) # добавляем каждый элемент в конец строки
    
            #добовляем значение
            worksheet['Z2'] = my_value

#<-% GROSS WEIGHT -end-------------
            last_row = worksheet.max_row
            # Сохраняем файл
            workbook.save(path)
            workbook.close()

#<-% NET WEIGHT --------------
'''
                
            
            #сканируем область       
            #gv = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(72, 0, 80, 94))        
            gv = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(81, 27, 85, 60) )       
            
            #проверка если место будет пустое 
            #CUSTOMER = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 4, 5, 6))        
            
            #делаем из отосланного массив
            send_CUST = np.array(gv)
    
            my_value = ''

            my_list = []

            # проход по элементам массива и добавление каждого элемента в строку
            for i in range(len(send_CUST)):
                for j in range(len(send_CUST[i])):
                    for k in range(len(send_CUST[i][j])):
                        my_value = str(send_CUST[i][j][k])
                        print (my_value)
                       
                        if str(send_y1[i][j][k]) == 'nan':
                            pass
                        else:
                            my_list.append(send_y1[i][j][k])

                        #print (send_CUST[i][j][k])
                        #my_value = str(send_CUST[i][j][k]) # добавляем каждый элемент в конец строки
                        
                        #просматриваем трайной массив и отправляем его по частям 
            if not my_list:
                pass
            else:
                if len(send_CUST[0][0]) == 3:
                    worksheet['AA2'] = send_CUST[0][0][1]
                    worksheet['AB2'] = send_CUST[0][0][2]
                if len(send_CUST[0][0]) == 2 and send_CUST[0][0][0] == my_value:
                    worksheet['AA2'] = send_CUST[0][0][1]
                    
                if len(send_CUST[0][0]) == 1 and send_CUST[0][0] == my_value:
                    pass
                else:
                    worksheet['AA2'] = send_CUST[0][0][0]
                            

'''   
            #добовляем значение
            #worksheet['AA2'] = my_value

#<-% NET WEIGHT -end-------------

#<-% VOLUME --------------
'''
                
            
            #сканируем область       
            #gv = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(72, 0, 80, 94))        
            vol = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(81, 27, 85, 60) )       
            
            #проверка если место будет пустое 
            #CUSTOMER = tabula.read_pdf(file, stream = True, multiple_tables = False, pages = page_num, relative_area = True,area=(3, 4, 5, 6))        
            
            #делаем из отосланного массив
            send_CUST = np.array(vol)
    
            my_value = ''

            # проход по элементам массива и добавление каждого элемента в строку
            for i in range(len(send_CUST)):
                for j in range(len(send_CUST[i])):
                    for k in range(len(send_CUST[i][j])):
                        #print (send_CUST[i][j][k])
                        my_value = str(send_CUST[i][j][k]) # добавляем каждый элемент в конец строки
    
            #добовляем значение
            #worksheet['AB2'] = my_value
'''
#<-% VOLUME -end-------------

#прр


#----------------------------main 



fromPDftoExelspec(file)
