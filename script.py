#!/usr/local/bin python3

import excel
import read_word
import openpyxl
import docx
import datetime
import os

name_of_excel = input('What is the name of the excel file?: ')
name_of_word = input('What is the name of the word document?: ')

list_of_rows = excel.read_excel(name_of_excel)
_range = len(list_of_rows)-1
word_doc_string = read_word.read_doc(name_of_word)
todays_date = datetime.datetime.now()
x = str(todays_date.strftime("%x"))

folder_name = input('What would you like to name the folder for the documents zachary?: ')

os.system('mkdir {}'.format(folder_name))

for iters in range(_range):
    new_doc_string = word_doc_string.format(A=list_of_rows[iters][0], B=list_of_rows[iters][1], C=list_of_rows[iters][2], D=list_of_rows[iters][3], E=x)
    new_doc_name = input('What would you like to call this word document?: ') +'.docx'
    new_doc = docx.Document()
    new_doc.add_paragraph(new_doc_string)
    new_doc.save('{}/{}'.format(folder_name, new_doc_name))