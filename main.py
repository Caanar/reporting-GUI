#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# Import libraries
import pandas as pd
import psycopg2
from sshtunnel import SSHTunnelForwarder
from tkinter import *
from tkinter import ttk
import os
from datetime import date
import xlsxwriter

# Establish SSH Tunnel connection and connect to Prod
tunnel = SSHTunnelForwarder(('sqlproxy.imsdev.io', 22), ssh_private_key='C:/Users/CArrieta/Desktop/id_rsa', ssh_username='carrieta', ssh_private_key_password='super_secret', remote_bind_address=('vread.work.net', 5432))
tunnel.start()
conn = psycopg2.connect(database='databasename',user='carrieta',host='localhost',password='secret',port= tunnel.local_bind_port,options="""-c search_path="public" """)
cur = conn.cursor()

# Root
root = Tk()
root.title('Customer Reporting - GUI')
root.geometry('700x300')

# Account Dropdown
label = Label(root, text='Account:')
label.pack()

cur.execute('''select name, id from accounts a where a.status in ('Active') order by name''')
colnames = [desc[0] for desc in cur.description]
rows = cur.fetchall()
df = pd.DataFrame(rows)
i = 0
for col in colnames:
    df.rename(columns={i: col},inplace=True)
    i+=1
options = df['name'].tolist()
combo1 = ttk.Combobox(root, value=options)
combo1.current(0)
combo1.config(width=60)
combo1.pack()
dict_to_acc_id = {}
for i in df.iterrows():
    dict_to_acc_id[i[1][0]] = i[1][1]

# Report type dropdown
label = Label(root, text='Report Type')
label.pack() # .grid(column=0, row=0) for grid format

options2 = os.listdir('C:/Users/CArrieta/Desktop/gui/queries')
combo2 = ttk.Combobox(root, value=options2)
combo2.current(0)
combo2.config(width=60)
combo2.pack()

# Field for zendesk email
label = Label(root, text='Email Verification')
label.pack()
textbox = Entry(root, width=40, bg='pink', borderwidth=5)
textbox.pack()
textbox.insert(0, 'Insert email')

# Button Function
def func():
    func_label = Label(root,text='')
    func_label2 = Label(root,text='')
    cur.execute(open(r"C:/Users/CArrieta/Desktop/gui/queries/{}".format(combo2.get())).read().replace(':account_id', str(dict_to_acc_id[combo1.get()])))
    colnames = [desc[0] for desc in cur.description]
    rows = cur.fetchall()
    df = pd.DataFrame(rows)
    i = 0
    for col in colnames:
        df.rename(columns={i: col},inplace=True)
        i+=1
    try:
        cur.execute('''select us.email, us.firstname, us.lastname from users us where us.account_id = {}'''.format(dict_to_acc_id[combo1.get()]))
        rows = cur.fetchall()
        admin_emails = pd.DataFrame(rows)
        admin_emails.set_index(0, inplace=True)
        if textbox.get() in admin_emails.index:
            func_label3 = Label(root,text='Success!!! That Email Exists in the selected account and belongs to {0} {1}'.format(admin_emails.loc[['{}'.format(textbox.get())]][1][0],admin_emails.loc[['{}'.format(textbox.get())]][2][0]))
            func_label3.pack()
            df.to_excel('C:/Users/CArrieta/Desktop/gui/excel/{0} Report {1}.xlsx'.format(combo1.get(),date.today()), sheet_name='{0}'.format(combo2.get()))
            writer = pd.ExcelWriter('C:/Users/CArrieta/Desktop/gui/excel/{0} Report {1}.xlsx'.format(combo1.get(),date.today()), engine='xlsxwriter')
            df.to_excel(writer, sheet_name='{0}'.format(combo2.get()), index=False)
            worksheet = writer.sheets['{0}'.format(combo2.get())]  # pull worksheet object
            for idx, col in enumerate(df):  # loop through all columns
                series = df[col]
                max_len = max((series.astype(str).map(len).max(), len(str(series.name)))) + 1  
                worksheet.set_column(idx, idx, max_len)
            writer.close()
            os.system('start "excel" "C:/Users/CArrieta/Desktop/gui/excel/{0} Report {1}.xlsx"'.format(combo1.get(),date.today()))            
        else:
            func_label4 = Label(root, text='Beware: that email is not used by any Admin Profile on that account')
            func_label4.pack()
    except:
        func_label2.config(text='Nope, no Excel for you')
        func_label2.pack()
    
button = Button(root, text='Get', padx= 50, command=func, bg='lightblue')
button.pack()

root.mainloop()


# In[ ]:




