import os
import re
import docx
import textract
import PyPDF2
import extract_msg
import xlrd
from docx import Document
from itertools import groupby
import pandas as pd
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords

#vars for this crawly boi
global PATTERN,output_path,error_path

path = ""
output_path = ""
error_path = ""
formats = ('.xlsx','.pdf','.doc','.txt','.docx','.xls','.msg')
#Arcane patterns for all types of credit cards
PATTERN1 = '^([3456][0-9]{3})[-]?([0-9]{4})[-]?([0-9]{4})[-]?([0-9]{4})$'

def count_consecutive(num):
    if max(len(list(g)) for _, g in groupby(num)) >= 4:
        return True
    else:
        return False

#This function checks if the arcane dark arts looking PATTERN variable is present in a line of text.
def is_cc_number(string, file_path, line_num,complex):
    string = string.replace(' ', '')
    if complex == False:
        match = re.match(PATTERN1, string)
    else:
        match = re.findall(PATTERN1, string)
    if not match or count_consecutive(string.replace('-', '')):
        pass
    else:
        output_file = open(output_path, "a+")
        output_file.write(f'\n{file_path}\nLine Number: {line_num}\nMatch: {match}\n',)
        output_file.close()



#Parse through the filtered items and see if there is any credit card numbers
def txt_check(files_to_check,path):
    print(' ')
    print('Parsing .txt files')
    for item in files_to_check:
        #Parse .txt files
        if '.txt' in item:
            try:
                print(f'Looking through: {path}/{item}')
                item = f'{path}/{item}'
    #            with open(item, "r") as f:
                line_num = 0
                with open(item, encoding="utf8", errors='ignore') as f:
                    for line in f:
                        line_num = line_num + 1
                        is_cc_number(line, f'{item}', line_num,complex=False)
            except:
                print('Skipping, .txt file is not supported.')
                pass


def docx_check(files_to_check,path):
    print(' ')
    print('Parsing .docx files')
    for item in files_to_check:
        #parse .docx files
        if  item.endswith('.docx'):
            print(f'Looking through: {path}/{item}')
            try:
                doc = docx.Document(f'{path}/{item}')
                line_num = 0
                for paragraph in doc.paragraphs:
                    line_num = line_num + 1
                    is_cc_number(paragraph.text, f'{path}/{item}', line_num, complex=False)
            except Exception as e:
                print('Skipping, .docx file is not supported.')
                error_output(f'{path}/{item}',e)
                pass




def doc_check(files_to_check,path):
    #Checks old ms word files
    print(' ')
    print('Parsing .doc files')
    for item in files_to_check:
        #parse .doc file
        if item.endswith('.doc'):
            print(f'Looking through: {path}/{item}\n')
            try:
                doc = textract.process(f"{path}/{item}")
                doc = doc.decode("utf-8")
                doc = doc.splitlines()
                line_num = 0
                for line in doc:
                    line_num = line_num + 1
                    is_cc_number(line,f'{path}/{item}', line_num, complex=False)
            except Exception as e:
                print('Skipping, .doc file is not supported.')
                error_output(f'{path}/{item}',e)
                pass



def pdf_check(files_to_check,path):
    print(' ')
    print('Parsing .pdf files')
    for item in files_to_check:
        #parse .pdf file
        if item.endswith('.pdf'):
            print(f'Looking through: {path}/{item}')
            try:
                filename = f'{path}/{item}'
                pdfFileObj = open(filename, 'rb')
                pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
                num_pages = pdfReader.numPages
                count = 0
                text = ""
                #Read each page
                while count < num_pages:
                    pageObj = pdfReader.getPage(count)
                    count += 1
                    text += pageObj.extractText()
                #Lets add a failsafe for scanned files
                if text != "":
                    text = text
                else:
                    text = textract.process(filename, method='tesseract', language='eng')
                #Time to clean the text variable
                #Do this if the file is actually a hidden html file
                try:
                    tokens = word_tokenize(text)
                except TypeError:
                    text = text.decode('utf-8')
                    tokens = word_tokenize(text)
                punctuations = ['(',')',';',':','[',']',',']
                stop_words = stopwords.words('english')
                keywords = [word for word in tokens if not word in stop_words and not word in punctuations]
                for word in keywords:
                    print(word)
                    is_cc_number(word,f'{path}/{item}','N/A', complex=False)
            except Exception as e:
                print('Skipping, .pdf file exception.')
                error_output(f'{path}/{item}',e)
                pass





def xlsx_check(files_to_check,path):
    print(' ')
    print('Parsing .xlsx files')
    for item in files_to_check:
        if item.endswith('.xlsx'):
            print(f'Looking through: {path}/{item}')
            try:
                doc = textract.process(f"{path}/{item}")
                doc = doc.decode("utf-8")
                doc = doc.splitlines()
                line_num = 0
                for line in doc:
                    line_num = line_num + 1
                    is_cc_number(line,f'{path}/{item}', line_num=line_num,complex=False)
            except Exception as e:
                print('Skipping, .xlsx file is not supported.')
                error_output(f'{path}/{item}',e)
                pass


def xls_check(files_to_check,path):
    print(' ')
    print('Parsing .xls files')
    for item in files_to_check:
        if item.endswith('.xls'):
            print(f'Looking through: {path}/{item}')
            #Create dataframe from first sheet
            try:
                xls_sheets = pd.read_excel(f'{path}/{item}', sheet_name=None)
                xls_sheets = str(xls_sheets)
                xls_sheets = xls_sheets.splitlines()
                line_num = 0
                for sheet in xls_sheets:
                    line_num = line_num + 1
                    line = sheet[2:]
                    is_cc_number(line,f'{path}/{item}', line_num=line_num,complex=False)
            except xlrd.XLRDError:
                try:
                    xls_sheets = pd.read_html(f'{path}/{item}')
                    xls_sheets = str(xls_sheets)
                    xls_sheets = xls_sheets.splitlines()
                    line_num = 0
                    for sheet in xls_sheets:
                        line_num = line_num + 1
                        line = sheet[2:]
                        is_cc_number(line,f'{path}/{item}', line_num=line_num,complex=False)
                except Exception as e:
                    print('Skipping, .xls file is not supported.')
                    error_output(f'{path}/{item}',e)
                    pass
            except Exception as e:
                print('Skipping, .xls file is not supported.')
                error_output(f'{path}/{item}',e)
                pass



def msg_check(files_to_check,path):
    print(' ')
    print('Parsing .msg files')
    for item in files_to_check:
        if item.endswith('.msg'):
            print(f'Looking through: {path}/{item}')
            try:
                msg = extract_msg.Message(f'{path}/{item}')
                msg = msg._getStringStream('__substg1.0_1000')
                msg = msg.splitlines()
                for line in msg:
                    is_cc_number(line,f'{path}/{item}', line_num='N/A', complex=False)
            except Exception as e:
                print('Skipping, .msg file is not supported.')
                error_output(f'{path}/{item}',e)
                pass


def check_all_file_types(files_to_check, dir):
    print(' ')
    print(f'FOUND DOCUMENTS IN {dir}:')
    if len(files_to_check) == 0:
        print('None')
    else:
        for item in files_to_check:
            print(item)
        txt_check(files_to_check,dir)
        xlsx_check(files_to_check,dir)
        xls_check(files_to_check,dir)
        msg_check(files_to_check,dir)
        pdf_check(files_to_check,dir)
        docx_check(files_to_check,dir)
        doc_check(files_to_check,dir)

def dir_crawl(path):
    print(f'Building directory tree for {path}')
    subdir_list = [x[0] for x in os.walk(path)]
    print('Building Complete')
    print(subdir_list)
    #Fetch the files to check for every subdirectory
    for dir in subdir_list:
        try:
            files_to_check = []
            path_contents = [x for x in os.listdir(dir)]
            for format in formats:
                for item in path_contents:
                    if format in item:
                        #Create a list of files for the directory and then start a check for each file in the list
                        files_to_check.append(item)
                        files_to_check = [ x for x in files_to_check if not x.startswith('~')]
            check_all_file_types(files_to_check, dir)
        except FileNotFoundError as e:
            error_output(dir,e)
    print('Done')

def error_output(dir,error):
    print(f'Error found\nError Type: {error}')
    file = open(error_path, "a+")
    file.write(f'\n{dir}\nError: {error}\n',)
    file.close()

###########################################################################################################################

dir_crawl(path)
