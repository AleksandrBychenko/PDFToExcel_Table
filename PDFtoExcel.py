import PyPDF2
import textract
import numpy as np
import openpyxl
import math
import os
import tabula
import pandas as pd

import tkinter as tk
from tkinter import filedialog
import PyPDF2
import PDFToTEXT3 as pd3
from PDFToTEXT3 import page_num
def browse_file():
    status_label.config(text="DO!")
    try:
        filepath = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        #status_label.config(text="DO")
        
        if filepath:
            #status_label.config(text=f"DO")
            pd3.fromPDftoExelspec(filepath)
        
            '''
            with open(filepath, 'rb') as pdf_file:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                num_pages = len(pdf_reader.pages)
                # do something with the pdf contents
            '''
            #status_label.config(text=f"Done!\n Take a special Pdf file to convert to an excel file ")
        status_label.config(text=f"Done!\n\n Take a special Pdf file to convert to an excel file.\n\n Remove the processed excel files\n from the folder where they appeared!!! ")


    except Exception:
        status_label.config(text="Error!!!\n" + "\nMost likely the file format is not executed.")
        # здесь можно добавить дополнительный код для обработки ошибки, если это необходимо
    




root = tk.Tk()
root.title("PDF Reader to exel")
root.geometry("300x250")
root.configure(bg='green')
#canvas = Canvas(root, width=ширина, height=высота)Browse

browse_button = tk.Button(root, text="Take PDF file", command = browse_file,width= '10',height= '5' )
browse_button.pack()

status_label = tk.Label(root, text="Take a special Pdf file to convert to an excel file.\n\n Remove the processed excel files\n from the folder where they appeared\n if you have already used this application !!!",width='40', height= '10')
status_label.pack()

root.mainloop()