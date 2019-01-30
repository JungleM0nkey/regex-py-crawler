import os
import re
import docx
import textract
import PyPDF2
import extract_msg
import xlrd
from docx import Document
global PATTERN,output_path
#Path for the start directory, works with windows and linux
path = "your/directory/here"
#Where the output file is saved
output_path = "your/directory/here"
formats = ('.xlsx','.pdf','.doc','.txt','.docx','.xls','.msg')
#Arcane looking regular expression goes here
PATTERN1 ='your regex pattern here'


#This function checks if the arcane dark arts looking PATTERN variable is present in a string.
def is_pattern_present(string, file_path, line_num):
    match = re.findall(PATTERN1, string)
    if match:
        output_file = open(output_path, "a+")
        output_file.write(f'\n{file_path}\nLine Number: {line_num}\nMatch: {match}\n',)
        output_file.close()
    else:
        pass


def txt_check(files_to_check,path):
    print(' ')
    print('Parsing .txt files')
    for item in files_to_check:
        #Parse .txt files
        if '.txt' in item:
            print(f'Looking through: {path}/{item}')
            item = f'{path}/{item}'
#            with open(item, "r") as f:
            line_num = 0
            with open(item, encoding="utf8", errors='ignore') as f:
                for line in f:
                    line_num = line_num + 1
                    is_pattern_present(line, f'{item}', line_num)


def docx_check(files_to_check,path):
    print(' ')
    print('Parsing .docx files')
    for item in files_to_check:
        #parse .docx files
        if '.docx' in item:
            print(f'Looking through: {path}/{item}')
            try:
                doc = docx.Document(f'{path}/{item}')
                line_num = 0
                for paragraph in doc.paragraphs:
                    line_num = line_num + 1
                    is_pattern_present(paragraph.text, f'{path}/{item}', line_num)
            #This is raised if there is nothing in the document, so we just ignore files that are empty
            except (docx.opc.exceptions.PackageNotFoundError, IsADirectoryError) as e:
                pass


def doc_check(files_to_check,path):
    #Checks old ms word files
    print(' ')
    print('Parsing .doc files')
    for item in files_to_check:
        #parse .doc file
        if item.endswith('.doc'):
            print(f'Looking through: {path}/{item}')
            try:
                doc = textract.process(f"{path}/{item}")
                doc = doc.decode("utf-8")
                is_pattern_present(doc,f'{path}/{item}', line_num='N/A')
            except IsADirectoryError:
                pass


def pdf_check(files_to_check,path):
    print(' ')
    print('Parsing .pdf files')
    for item in files_to_check:
        #parse .pdf file
        if item.endswith('.pdf'):
            print(f'Looking through: {path}/{item}')
            try:
                doc = open(f'{path}/{item}', 'rb')
                pdfreader = PyPDF2.PdfFileReader(doc)
                total_pages = pdfreader.getNumPages()
                for page_number in range(total_pages):
                    page = pdfreader.getPage(page_number)
                    page_content = page.extractText()
                    print(f'Looking through {item} PDF Page: {page_number+1}')
                    is_pattern_present(page_content,f'{path}/{item}',page_number+1)
            except IsADirectoryError:
                pass


def xlsx_check(files_to_check,path):
    print(' ')
    print('Parsing .xlsx files')
    for item in files_to_check:
        #parse .doc file
        if item.endswith('.xlsx'):
            print(f'Looking through: {path}/{item}')
            try:
                doc = textract.process(f"{path}/{item}")
                doc = doc.decode("utf-8")
                is_pattern_present(doc,f'{path}/{item}', line_num='N/A')
            except IsADirectoryError:
                pass

def xls_check(files_to_check,path):
    print(' ')
    print('Parsing .xls files')
    for item in files_to_check:
        #parse .doc file
        if item.endswith('.xls'):
            print(f'Looking through: {path}/{item}')
            try:
                doc = textract.process(f"{path}/{item}")
                doc = doc.decode("utf-8")
                is_pattern_present(doc,f'{path}/{item}', line_num='N/A')
            except xlrd.XLRDError:
                line_num = 0
                with open(f"{path}/{item}", encoding="ISO-8859-1") as f:
                    for line in f:
                        line_num = line_num + 1
                        is_pattern_present(line, f'{path}/{item}', line_num='N/A')
            except IsADirectoryError:
                pass


def msg_check(files_to_check,path):
    print(' ')
    print('Parsing .msg files')
    for item in files_to_check:
        #parse .doc file
        if item.endswith('.msg'):
            print(f'Looking through: {path}/{item}')
            msg = extract_msg.Message(f'{path}/{item}')
            msg = msg._getStringStream('__substg1.0_1000')
            is_pattern_present(msg,f'{path}/{item}', line_num='N/A')
            

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

		
#The recursive crawl through directories function, pretty straight forward.
def dir_crawl(path):
    subdir_list = [x[0] for x in os.walk(path)]
    #Fetch the files to check for every subdirectory
    for dir in subdir_list:
        files_to_check = []
        path_contents = [x for x in os.listdir(dir)]
        for format in formats:
            for item in path_contents:
                if format in item:
                    #Create a list of files for the directory and then start a check for each file in the list
                    files_to_check.append(item)
					#Do not look through temp files, it breaks things.
                    files_to_check = [ x for x in files_to_check if not x.startswith('~')]
        check_all_file_types(files_to_check, dir)

###########################################################################################################################

dir_crawl(path)
