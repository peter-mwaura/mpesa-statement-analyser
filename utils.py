import pikepdf
from copy import copy
import PyPDF2
import io
import re
from io import StringIO, BytesIO
import re
from PyPDF2 import PdfFileReader
from openpyxl import Workbook
from openpyxl.styles import Font, Color, colors
import random, string
from openpyxl.writer.excel import save_virtual_workbook
import openpyxl
from openpyxl import load_workbook
import os
import pandas as pd
from pandas import Series
import pandas.io.formats.excel
pandas.io.formats.excel.header_style = None
import xlsxwriter
import pdfplumber
from datetime import datetime

i = 0
#regular expression for mpesa statement
#regex1 = r'(C.+ Name)(.+)(M.+ Number)(\d+)(E\w+ Address)(.+)(D[ a-zA-Z]+Statement)(.+)(S.+ Period)(.+ - \d{2} \w+ \d{4})'
#regex = r'(\w{10})(\d{4}-\d{2}-\d{2} \d{2}\:\d{2}\:\d{2})(.+?)(Completed)(.*?\.\d{2})(.*?\.\d{2})'


def extract_from_pdf(file, password):
    """
    1.Decrypts the file if encrypted with the pikepdf module
    2.Open the now decrypted file and stores the read text in a StringIO
      class object stored in the second variable which can be extracted 
      using the getvalue method
    3.Returns the number of pages and text content - Use the getvalue method
      to print the contents
    """
    #decrypting the encrypted pdf file
    content = pikepdf.open(file, password=password)
    print('content done')
    inmemory_file = BytesIO()
    content.save(inmemory_file)
    print('saved')
    #reading and extracting data from the decrypted pdf file 
    pdf_reader = pdfplumber.open(inmemory_file)
    print('read')
    num_pages = pdf_reader.pages
    print('getpages')
    #num_pages = 6
    #pages = pdf_reader.pages[0]
    #print(len(num_pages))

    extracted_data = StringIO()
    #extracted_data.writelines(pages.extract_text())
    
    now = datetime.now()
    for page in num_pages:
        #print(page, flush=True)
        extracted_data.writelines(page.extract_text())
        diff = datetime.now() - now
        print(diff)
        now = datetime.now()

    return num_pages, extracted_data


def parse_mpesa_content(extracted_data):
    new_regex_name = r'(C.+ Name: )(.*)(\n)(M.+ Number: )(\d+)(\n)(E\w+ Address)(.*)(\n)(S.+ Period: )(.*)(\n)(R.+Date: )(.*)'
    #new_regex_transactions = r'(\w{10} )(\d{4}-\d{2}-\d{2} \d{2}\:\d{2}\:\d{2})(.+?)(Completed )(.*?\.\d{2} )(.*?\.\d{2})'
    new_regex_transactions = r'(\w{10} )(\d{4}-\d{2}-\d{2} \d{2}\:\d{2}\:\d{2})(.+?)(Completed )(.*?\.\d{2} )(.*?\.\d{2})(.*?(?=(\w{10} \d{4}-\d{2}-\d{2})|Disclaimer:))'
    extracted_data.seek(0)
    lines = extracted_data.read()
    #matches = re.compile(regex)cd .findall(lines)
    #matches2 = re.compile(regex1).findall(lines)
    matches_name = re.compile(new_regex_name).findall(lines)
    matches_transactions = re.compile(new_regex_transactions, re.DOTALL).findall(lines)

    fb = Font(name='Calibri', color=colors.BLACK, bold=True, size=11, underline='single')
    i = 0
    #creating the spreadheet
    book = Workbook()
    # grab the active worksheet
    sheet = book.active
    #excel styling 2
    ft = Font(name='Calibri', color=colors.BLUE, bold=True, size=11, underline='single')

    sheet['A1'] = 'RECEIPT NO'
    sheet['B1'] = 'COMPLETION TIME'
    sheet['C1'] = 'DETAILS'
    sheet['D1'] = 'TRANSACTION STATUS'
    sheet['E1'] = 'VALUE'
    sheet['F1'] = 'BALANCE'

    a1 = sheet['A1']
    b1 = sheet['B1']
    c1 = sheet['C1']
    d1 = sheet['D1']
    e1 = sheet['E1']
    f1 = sheet['F1']

    a1.font = ft
    b1.font = ft
    c1.font = ft
    d1.font = ft
    e1.font = ft
    f1.font = ft


    #adding every match to the excel file
    #while i < len(matches):
        # print(matches[i])
        #sheet.append(matches[i])
        #i = i + 1

    for match in matches_transactions:
        match = list(match)
        match[2] = match[2] + match[6]
        sheet.append(match[:-2])

    filename = random_str() + '.' + 'xlsx'
    book.save(filename)
    f = open(filename, 'rb')
    file = BytesIO(f.read())
    f.close()
    os.remove(filename)

    return file, matches_name[0][1]

def find_name(matches2):
    for match in matches2:
        print(match[1])

    return match[1]

def random_str(length=8):
    s = ''
    for i in range(length):
        s += random.choice(string.ascii_letters + string.digits)

    return s

def summary(workbook):
    excel_df = pd.read_excel(workbook)
    excel_df['VALUE'] = excel_df['VALUE'].astype(str).str.replace(',', '').astype(float)
    excel_df['COMPLETION TIME'] = pd.to_datetime(excel_df['COMPLETION TIME'])
    excel_df['month_of_date'] = excel_df['COMPLETION TIME'].dt.month

    paidinall = excel_df[excel_df['VALUE']>0]
    withdrawn = excel_df[excel_df['VALUE']<0]

    paidin = paidinall[['month_of_date', 'VALUE']].groupby(['month_of_date'], as_index=False)['VALUE'].sum()
    paidin.rename(columns={'VALUE':'Paid In'}, inplace=True)
    paidtotal = paidin['Paid In'].sum()

    withdraw = withdrawn[['month_of_date', 'VALUE']].groupby(['month_of_date'], as_index=False)['VALUE'].sum()
    withdraw.rename(columns={'VALUE':'Withdrawn'}, inplace=True)
    withdraw['Withdrawn'] = withdraw['Withdrawn'].astype(str).str.replace('-', '').astype(float)
    withdrawntotal = withdraw['Withdrawn'].sum()

    combined = pd.merge(paidin,withdraw,on='month_of_date')
    excel_df = pd.DataFrame({'month_of_date': 'Grand Total', 'Paid In':[paidtotal], 'Withdrawn':[withdrawntotal]})
    df_append = combined.append(excel_df, ignore_index=False)

    return df_append

def paidin(workbook):
    excel_df = pd.read_excel(workbook)
    excel_df['VALUE'] = excel_df['VALUE'].astype(str).str.replace(',', '').astype(float)
    paidinall = excel_df[excel_df['VALUE']>0]
    # paidinall.set_index('DETAILS', inplace=True)
    paidin = paidinall[['VALUE', 'DETAILS']].sort_values('DETAILS').groupby(['DETAILS'], as_index=False)['VALUE'].sum()
    def format(row):
        index = None
        reg = re.search(r'\d', row['DETAILS'])
        if reg:
            index = reg.start()
        row['DETAILS'] = row['DETAILS'][:index]
        return row

    sorted_df = paidin.apply(format, axis=1).groupby(['DETAILS'], as_index=False).apply(lambda r: r).sort_values(['DETAILS', 'VALUE'], ascending=False)
    idx = sorted_df.index
    paidin = paidin.loc[idx]

    unique_groups = set(sorted_df['DETAILS'])
    details_series = sorted_df['DETAILS']
    index_for_groups = {group: idx.get_loc(details_series.where(details_series==group).last_valid_index())
                        for group in unique_groups}

    values = sorted(index_for_groups.values())

    added = 0
    paidin = paidin.append(Series([]), ignore_index=True)
    for index in values:
        index += added
        paidin = paidin.loc[:index].append(Series([]), ignore_index=True).append(paidin.loc[index+1:], ignore_index=True)
        added += 1

    subtotal = paidin['VALUE'].sum()
    excel_df = pd.DataFrame({'VALUE':[subtotal], 'DETAILS': 'Grand Total'})
    df_append = paidin.append(excel_df, ignore_index=False)
    df_append.rename(columns={'VALUE':'AMOUNT'}, inplace=True)

    return df_append

def withdrawal(workbook):
    excel_df = pd.read_excel(workbook)
    excel_df['VALUE'] = excel_df['VALUE'].astype(str).str.replace(',', '').astype(float)
    withdrawal = excel_df[excel_df['VALUE']<0]
    withdrawn = withdrawal[['VALUE', 'DETAILS']].sort_values('DETAILS').groupby(['DETAILS'], as_index=False)['VALUE'].sum()
    withdrawn['VALUE'] = withdrawn['VALUE'].astype(str).str.replace('-', '').astype(float)
    def format(row):
        index = None
        reg = re.search(r'\d', row['DETAILS'])
        if reg:
            index = reg.start()
        row['DETAILS'] = row['DETAILS'][:index]
        return row

    sorted_df = withdrawn.apply(format, axis=1).groupby(['DETAILS'], as_index=False).apply(lambda r: r).sort_values(['DETAILS', 'VALUE'], ascending=False)
    idx = sorted_df.index
    withdrawn = withdrawn.loc[idx]

    unique_groups = set(sorted_df['DETAILS'])
    details_series = sorted_df['DETAILS']
    index_for_groups = {group: idx.get_loc(details_series.where(details_series==group).last_valid_index())
                        for group in unique_groups}

    values = sorted(index_for_groups.values())

    added = 0
    withdrawn = withdrawn.append(Series([]), ignore_index=True)
    for index in values:
        index += added
        withdrawn = withdrawn.loc[:index].append(Series([]), ignore_index=True).append(withdrawn.loc[index+1:], ignore_index=True)
        added += 1
    subtotal = withdrawn['VALUE'].sum()
    excel_df = pd.DataFrame({'VALUE':[subtotal], 'DETAILS': 'Grand Total'})
    df_append = withdrawn.append(excel_df, ignore_index=False)
    df_append.rename(columns={'VALUE':'AMOUNT'}, inplace=True)

    return df_append

def listing(summary, paidin, withdrawn):
    df = [summary, paidin, withdrawn]

    return df


def dfs_tabs(df_list, sheet_list, file_name):
    file_name = BytesIO()
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    workbook = writer.book
    fmt = workbook.add_format({'align':'left', 'size':10, 'font_name': 'Times New Roman'})
    for dataframe, sheet in zip(df_list, sheet_list):
        dataframe.to_excel(writer, sheet_name=sheet, startrow=0, startcol=0, index=False)
    worksheet1 = writer.sheets['SUMMARY']
    worksheet2 = writer.sheets['PAID IN DATA']
    worksheet3 = writer.sheets['WITHDRAWN DATA']
    worksheet1.set_column(0, 2, 40.0, fmt)
    worksheet2.set_column(0, 2, 90.0, fmt)
    worksheet3.set_column(0, 2, 90.0, fmt)
    writer.save()

    return file_name
