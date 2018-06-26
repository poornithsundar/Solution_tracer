import xlrd
import os
from tkinter import *
from tkinter import ttk

z = os.getcwd()
workbook = xlrd.open_workbook(z+'\\Output.xlsx')
sheets = workbook.sheet_names()
required_data = []
for sheet_name in sheets:
    sh = workbook.sheet_by_name(sheet_name)
    for rownum in range(sh.nrows):
        row_values = sh.row_values(rownum)
        required_data.append((row_values[0]))
required_data2 = []
for sheet_name in sheets:
    sh = workbook.sheet_by_name(sheet_name)
    for rownum in range(sh.nrows):
        row_values = sh.row_values(rownum)
        required_data2.append((row_values[1]))
required_data1 = list(filter(None, required_data))
top = Tk()
v = Label(top, text="Problem:")
v.pack()
z = os.getcwd()
workbook = xlrd.open_workbook(z+'\\Output.xlsx')
sheets = workbook.sheet_names()
required_data = []
for sheet_name in sheets:
    sh = workbook.sheet_by_name(sheet_name)
    for rownum in range(sh.nrows):
        row_values = sh.row_values(rownum)
        required_data.append((row_values[0]))
required_data1 = list(filter(None, required_data))
def on_keyrelease(event):

    # get text from entry
    value = event.widget.get()
    value = value.strip().lower()

    # get data from required_data1
    if value == '':
        data = required_data1
    else:
        data = []
        for item in required_data1:
            if value in item.lower():
                data.append(item)                

    # update data in listbox
    listbox_update(data)


def listbox_update(data):
    # delete previous data
    listbox.delete(0, 'end')

    # sorting data 
    data = sorted(data, key=str.lower)

    # put new data
    for item in data:
        listbox.insert('end', item)


def on_select(event):
    global ert
    ert = event.widget.get(event.widget.curselection())
    entry.delete(0,END)
    entry.insert(0,ert)


# --- main ---

entry = ttk.Entry(top, width=100)
entry.pack()
entry.bind('<KeyRelease>', on_keyrelease)

listbox = Listbox(top, width=100)
listbox.pack()
listbox.bind('<<ListboxSelect>>', on_select)
listbox_update(required_data1)
def z():
    y.delete('1.0',END)
def close_window ():
    z = entry.get()
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
w = ttk.Button(top, text="CHECK",command = close_window)
w.pack()
ww = ttk.Button(top, text="CLEAR",command = z)
ww.pack()
f = ttk.Scrollbar(top)
y = Text(top, yscrollcommand=f.set)
y.pack(side = LEFT)
f.pack(side = RIGHT, fill=Y)
f.config(command = y.yview)
def zz():
    top = Tk()
    v = Label(top, text="Your Gmail address:")
    v.pack()
    st = Entry(top)
    st.pack()
    v1 = Label(top, text="Your Gmail Password:")
    v1.pack()
    st1 = Entry(top,show="*")
    st1.pack()
    v2 = Label(top, text="Query:")
    v2.pack()
    st2 = Entry(top)
    st2.pack()
    def z():
        import smtplib
        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls()
        s.login(st.get(), st1.get())
        message = st2.get()
        s.sendmail(st.get(), "thotapoornithsundar@gmail.com", message)
        s.quit()
        top.destroy()
        y.insert(END,"Query Sent.....")
    v = ttk.Button(top, text = "LOGIN", command = z)
    v.pack()
    top.geometry("500x300")
    top.title("Email Feedback")
    top.mainloop()
ww = ttk.Button(top, text="Problem not listed?",command = zz)
ww.pack(side=BOTTOM)
top.title("ABC Corp - Solution Tracer")
top.mainloop()
