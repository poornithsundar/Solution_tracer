import xlrd 
import pandas as pd
import textwrap
import array
from tkinter import *
from tkinter import ttk
import time
import xlrd
import os
z = os.getcwd()
workbook = xlrd.open_workbook(z+'\\SampleInput.xlsx')
sheets = workbook.sheet_names()
required_data = []
for sheet_name in sheets:
    sh = workbook.sheet_by_name(sheet_name)
    for rownum in range(sh.nrows):
        row_values = sh.row_values(rownum)
        required_data.append((row_values[4]))
required_data2 = []
for sheet_name in sheets:
    sh = workbook.sheet_by_name(sheet_name)
    for rownum in range(sh.nrows):
        row_values = sh.row_values(rownum)
        required_data2.append((row_values[5]))        
required_data1 = list(filter(None, required_data))
top = Tk()
f = Scrollbar(top)
v = Label(top, text="Previous tickets:")
x = ttk.Combobox(top, width=90)
x['values']=(list(set(required_data1[1:])))
def z():
    y.delete('1.0',END)
def close_window ():
    z = x.get()
    for sheet in workbook.sheets():
        for rowidx in range(sheet.nrows):
            row = sheet.row(rowidx)
            for colidx, cell in enumerate(row):
                if cell.value == z :
                    aov = sheet.cell_value(rowidx,colidx+1)
                    aov1 = aov.split("<br />")
                    a = len(aov1)
                    lst=[]
                    for i in range(0,a):
                        lst.append(len(aov1[i]))
                    max_val = lst[0]       
                    for item in lst:       
                        if item > max_val: 
                            max_val = item
                    for i in range(0,a):
                        if lst[i] == max_val:
                            global y
                            y.insert(END,aov1[i]+"\n")
w = Button(top, text="SUBMIT",command = close_window)
ww = Button(top, text="CLEAR",command = z)
top.destroy()

import docx2txt
from textblob import TextBlob
from nltk import BlanklineTokenizer
import re
import tkinter as tk  
from tkinter import ttk
from tkinter import *

win = tk.Tk()
win.title("Python GUI App")  
ttk.Label(win, text="Choose the question:")
def z():
    y.delete('1.0',END)
def click():
    c = numberChosen.get()
    for i in range (0,len(z)):
        x = z[i].find(c)
        if x!= -1:
            global a
            a = z[i+1]
            y.insert(END,a)
action = ttk.Button(win, text="Click", command=click)  
ww = ttk.Button(win, text="Clear",command = z)
number= tk.StringVar()  
numberChosen= ttk.Combobox(win, width=50)
y = Text(win, height=10, width=100)
import os
z = os.getcwd()
text = docx2txt.process(z+"\\SampleInputDoc1-FAQs.docx")
blob = TextBlob(text)
tokenizer = BlanklineTokenizer()
z = blob.tokenize(tokenizer)
c = '?'
lst = list()
for i in range (0,len(z)):
    x = z[i].find(c)
    if x!= -1:
        lst.append(z[i])
numberChosen['values']=(lst)
win.destroy()

import xlsxwriter

workbook = xlsxwriter.Workbook('SampleOutput.xlsx')
worksheet = workbook.add_worksheet()


# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

# Iterate over the data and write it out row by row.
for problem in (required_data):
    worksheet.write(row, col, problem)
    row += 1

row = 0
col = 1

for solution in (required_data2):
    worksheet.write(row, col, solution)
    row += 1


row = 400
col = 0

# Iterate over the data and write it out row by row.
for i in range (0,len(z)):
    if i%2 == 0:
        worksheet.write(row, col, z[i])
        row += 1

row = 400
col = 1

for i in range (0,len(z)):
    if i%2 != 0:
        worksheet.write(row, col, z[i])
        row += 1

workbook.close()

print("Sample Output Generated,.....")
time.sleep(2)
