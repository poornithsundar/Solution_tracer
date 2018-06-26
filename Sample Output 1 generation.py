import time
import mammoth
import os
z = os.getcwd()
f = open(z+'\\SampleDocuments\\SampleInputDoc2-.docx','rb')
b = open('x.html','wb')
document = mammoth.convert_to_html(f)
b.write(document.value.encode('utf8'))
f.close()
b.close()

import codecs
r=codecs.open("x.html", 'r').read()

from bs4 import BeautifulSoup
import re
soup = BeautifulSoup(r,"lxml")
company_name = soup.find_all('strong')
company_name1 = soup.find_all('h3')

import re

import tkinter as tk
from tkinter import ttk
from tkinter import *
wert = tk.Tk()
z = ttk.Label (wert, text="Problem:")
f = Scrollbar(wert)
numberChosen= ttk.Combobox(wert, width=50)
def ze():
    y.delete('1.0',END)
def z():
    c = numberChosen.get()
    if c == "PC Won't Boot Into wertdows":
        for i in range (1,5):
            global a
            a = company_name[i]
            a = re.sub("<.*?>", "", a.text)
            y.insert(END,a)
            y.insert(END,'\n')
    if c == "PC Is Running Slowly":
        for i in range (5,12):
            a = company_name[i]
            a = re.sub("<.*?>", "", a.text)
            y.insert(END,a)
            y.insert(END,'\n')
    if c == "Internet Is Slow or Not Loading":
        for i in range (12,16):
            a = company_name[i]
            a = re.sub("<.*?>", "", a.text)
            y.insert(END,a)
            y.insert(END,'\n')
    if c == "Windows Explorer Is Hanging":
        a = "If you're having problems loading up Windows Explorer and browsing your file system, the problem is almost always a shell extension that shouldn't be installed, or some shell extensions that are conflicting with each other. For example, the shell extensions for Dropbox and TortoiseSVN tend to cause problems when you put your code into your Dropbox folder, causing hanging and generally slow file browsing.Your best bet is to grab a copy of ShellExView and start disabling third-party shell extensions, or uninstalling Windows Explorer plug-ins that you don't actually need. You can also use this tool in combination with ShellMenuView to clean up your messy Explorer context menu."
        y.insert(END,a)
lst = list()
asd = company_name[0]
a = re.sub("<.*?>", "", asd.text)
lst.append(a)
for i in range (0,len(company_name1)):
    x = company_name1[i]
    a = re.sub("<.*?>", "", x.text)
    lst.append(a)
numberChosen['values']=(lst)
y = Text(wert, height=10, width=100)
ww = ttk.Button(wert, text="Submit",command = z)
www = ttk.Button(wert, text="Clear",command = ze)
wert.destroy()

ad = []
for i in range (1,16):
            a = company_name[i]
            a = re.sub("<.*?>", "", a.text)
            ad.append(a)
az = "If you're having problems loading up Windows Explorer and browsing your file system, the problem is almost always a shell extension that shouldn't be installed, or some shell extensions that are conflicting with each other. For example, the shell extensions for Dropbox and TortoiseSVN tend to cause problems when you put your code into your Dropbox folder, causing hanging and generally slow file browsing.Your best bet is to grab a copy of ShellExView and start disabling third-party shell extensions, or uninstalling Windows Explorer plug-ins that you don't actually need. You can also use this tool in combination with ShellMenuView to clean up your messy Explorer context menu."
ad.append(az)
import xlsxwriter

workbook = xlsxwriter.Workbook('SampleOutput1.xlsx')
worksheet = workbook.add_worksheet()


# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

# Iterate over the data and write it out row by row.
for i in range (0,4):
    worksheet.write(row, col, lst[i])
    row += 1

row = 0
col = 1

worksheet.write(row, col, ad[0]+". "+ad[1]+". "+ad[2]+"."+ad[3]+".")
row+=1

worksheet.write(row, col, ad[4]+". "+ad[5]+". "+ad[6]+"."+ad[7]+"."+ad[8]+". "+ad[9]+"."+ad[10]+".")
row+=1

worksheet.write(row, col, ad[11]+". "+ad[12]+". "+ad[13]+"."+ad[14]+".")
row+=1

worksheet.write(row, col, az)
workbook.close()

print("Sample Output 1 Generated,.....")
time.sleep(2)
