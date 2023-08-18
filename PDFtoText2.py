import PyPDF2
import textract
import numpy as np
import openpyxl

# Открываем PDF-файл
with open("54564-54566.pdf", 'rb') as file:
    # Создаем объект для чтения PDF-файла
    pdf_reader = PyPDF2.PdfReader(file)

    # Создаем массив для хранения текстовых данных
    text_data = []

    # Открываем Excel-файл
    workbook = openpyxl.Workbook()
    #workbook = openpyxl.load_workbook('file.xlsx')

    # Проходим по каждой странице PDF-файла
    for page_num in range(len(pdf_reader.pages)):
        # Получаем объект страницы
        page = pdf_reader.pages[page_num]

        # Извлекаем текстовые данные из страницы
        text = page.extract_text()

        # Делим текст на строки и сохраняем в массив
        lines = text.split('\n')
        

        #_______________>
        '''
        #считываем строчку 
        send = []

        send1 = ['CUSTOMER']
        send.append(send1)

        send2  = []
        import re
        ln1 = re.split('\s+', lines[1])
        send2.append(str(ln1[4])+ ' '+str(ln1[5]) + '\n' + str(lines[2]) + '\n' + str(lines[3]) + '\n' + str(lines[4]))
        send.append(send2)

        send3  = ['Cust Code', 'VAT NO', 'NO', 'Date', 'Page']
        send.append(send3)

        import re
        #send4 = re.split('\s+', lines[7])
        #print(lines[0], lines[1], lines[2], lines[3],lines[4], lines[5], lines[6],lines[7])
        send4 = []
        ln1 = re.split('\s+', lines[1])
        print(lines[6].split(' ')[0], ln1, lines[0].split(' '))
        send4.append(lines[0].split(' ')[0])
        send4.append(lines[6].split(' ')[0])
        send4.append(ln1[1])
        send4.append(ln1[2])
        send4.append(ln1[3])
        
        #send.append(send4)
        #print(send)


        # Получаем активный лист
        worksheet = workbook.active

        # Создаем массив для новой строки
        new_row = ['Значение 1', 'Значение 2', 'Значение 3']

        # Добавляем новую строку в конец листа
        #for i in range(len(send)):
        #    worksheet.append(send[i])
        #   workbook.save('file.xlsx')
        

        worksheet.append(send1)
        worksheet.append(send2)

        worksheet.append(send3)
        
 
        worksheet.append(send4)
        # Сохраняем изменения в файл
        workbook.save('file.xlsx')
        '''

        # Добавляем текстовые данные в массив
        
        #text_data.append(send)

        
# Выводим массив с текстовыми данными
#print(text_data)