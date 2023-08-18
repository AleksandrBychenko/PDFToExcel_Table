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
        
        #считываем строчку 
        send  = []
        send.append(str(lines[1]) + '\n' + str(lines[2]) + '\n' + str(lines[3]) + '\n' + str(lines[4]))
    
        print(send)


        # Получаем активный лист
        worksheet = workbook.active

        # Создаем массив для новой строки
        new_row = ['Значение 1', 'Значение 2', 'Значение 3']

        # Добавляем новую строку в конец листа
        worksheet.append(send)

        # Сохраняем изменения в файл
        workbook.save('file.xlsx')

        # Добавляем текстовые данные в массив
        text_data.append(send)

# Выводим массив с текстовыми данными
#print(text_data)