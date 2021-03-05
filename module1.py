# Preparing timings Calculations By Swipe Data Form of Excel or Csv Format files
# Version 1.0
# Prepared By : G.Devi Krishna Varma
# Module Task : Open Excel or Csv  & Read & Calculations
# Module Year : March 2020
# Copyrights  : Apache(By Git hub)
# copyright@krishnavarama
from tkinter import *
import os, sys,socket,os,time
from pandastable import Table
from tkinter.filedialog import askopenfilename
import pandas as pd
import numpy as np

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


root = Tk()
sw = root.winfo_screenwidth()
sh = root.winfo_screenheight()
sysn = socket.gethostname()
root.geometry('%sx%s+%s+%s' % (sw, sh, 0, 0))
f = int(sw / 2)
s = '%s+%s' % (f, 0)
root.title(" EMPLOYEE WORKING HOURS ")
p1 = PhotoImage(file=resource_path('clock.png'))
root.iconphoto(False, p1)


def del_des():
    aot.destroy()
    data_frm()
    if os.path.exists(t_path):
        os.remove(t_path+'pivot.csv')
    if os.path.exists(t_path):
        os.remove(t_path+'sum.csv')
    if os.path.exists(t_path+'Timedata.csv'):
        os.remove(t_path+'Timedata.csv')
    else:
        pass

def t_html():
    if br_file is not None:
        df.to_html(f_path+"working_hours.html")
        e.to_html(f_path+'Working_Hours_By_Date.html')
        l1 = Label(root, text=" Sucessfully coverted to HTML Files ", font='Arial 12 bold', fg='green')
        l1.pack(ipady=5,fill='x')
        l1.after(1000, lambda: l1.destroy())

def t_excel():
    if br_file is not None:
        df.to_excel(f_path+"working_hours.xlsx")
        e.to_excel(f_path+'Working_Hours_By_Date.xlsx')
        l2 = Label(root, text=" Sucessfully coverted to Excel Files ", font='Arial 12 bold', fg='green')
        l2.pack(ipady=5,fill='x')
        l2.after(1000, lambda: l2.destroy())

def t_csv():
    if br_file is not None:
        df.to_csv(f_path+"working_hours.csv")
        e.to_csv(f_path+'Working_Hours_By_Date.csv')
        l3 = Label(root, text=" Sucessfully coverted to CSV Files ", font='Arial 12 bold', fg='green')
        l3.pack(ipady=2,fill='x')
        l3.after(1000, lambda: l3.destroy())

def by_date():
    global g
    global aot
    global e
    if br_file is not None:
        g = Button(home, text='Back', bg='orange', height=1, width=5, command=lambda: [del_des(), g.destroy()])
        g.pack(pady=2,ipadx=15,padx=8,anchor=NE)
    ptf.destroy()
    d = pd.DataFrame()
    d['Card Number'] = df['Card Number']
    d['Date'] = df['Date']
    d['A'] = df['InTime']
    d['B'] = df['OutTime']

    fmt = '%H:%M:%S'

    d['Ins'] = pd.to_datetime(d['A'],
                              format=fmt,
                              errors='coerce')
    d['Outs'] = pd.to_datetime(d['B'],
                               format=fmt,
                               errors='coerce')

    d['Spending Times'] = d['Ins'] - d['Outs']
    d['Spending Times'] = d['Spending Times'] / np.timedelta64(1, 's')
    e = d.groupby(['Card Number', 'Date'], as_index=False).sum()
    e.to_csv(t_path+"sum.csv", index=False)  # 1
    su = pd.read_csv(t_path+"sum.csv", index_col=False)
    su['seconds'] = su['Spending Times'].abs().round(decimals=1).astype(int)
    seconds = su['seconds'] % (24 * 3600)
    su['hours'] = su['seconds'] // 3600
    su['seconds'] %= 3600
    su['minutes'] = su['seconds'] // 60
    su['seconds'] %= 60
    su['Day'] = su['hours'].astype(str) + ':' + su['minutes'].astype(str) + ':' + su['seconds'].astype(str)
    su['instamp'] = pd.to_datetime(su['Day'], format=fmt,
                                   errors='coerce')
    su['Day'] = su['instamp'].dt.time
    su.drop(['hours', 'minutes', 'seconds', 'instamp', 'Spending Times'], inplace=True, axis=1)
    e = su.pivot('Card Number', 'Date').rename_axis().fillna('A')
    e.to_csv(t_path+'pivot.csv')
    dp = pd.read_csv(t_path+"pivot.csv")
    aot = Frame(home)
    aot.pack(fill=BOTH, expand=1)
    table1 = pto1 = Table(aot, dataframe=dp,
                         showtoolbar=False, showstatusbar=False,editable=False,align='center')
    table1.autoResizeColumns()
    pto1.show()


def b_file():
    global d_path
    global t_path
    global f_path
    dt= time.strftime('%d-%m-%y')
    t_path = 'C:\Windows\Temp/'#os.environ['HOMEPATH']+'\Desktop\EWH_Output_Files\%s\temp'%(dt)
    d_path = os.environ['HOMEPATH']+'\Desktop\EWH_Output_Files'
    if not os.path.exists(d_path):
        os.makedirs(d_path)
    f_path = os.environ['HOMEPATH']+'\Desktop\EWH_Output_Files\%s/'%(dt)
    if not os.path.exists(f_path):
        os.makedirs(f_path)
        
    f_path = os.environ['HOMEPATH']+'\Desktop\EWH_Output_Files\%s/'%(dt)

    #os.makedirs(t_path)
    global br_file
    br_file = askopenfilename()
    data_frm()


def data_frm():
    global ptf
    global br_file
    global df
    if br_file is not None:
        try:
            df = pd.read_csv("%s" % (br_file), index_col=False,dtype=object)
            #df = pd.read_csv(t_path+"Timedata.csv", index_col=False)
            #fnm = os.rename('%s' % (br_file), 'Timedata.csv')
        except:
            df = pd.read_excel("%s" % (br_file))
            df.to_csv(t_path+"Timedata.csv", index=False)
            df = pd.read_csv(t_path+"Timedata.csv", index_col=False,dtype=object)

        df['timestamp'] = pd.to_datetime(df['Event Date'])
        df['Date'] = df['timestamp'].dt.date
        df['In_Time'] = df['timestamp'].dt.time
        df.drop_duplicates(keep='first', inplace=True)
        start, stop, step = 24, 27, 1
        df["inout"] = df["Location"].str.slice(start, stop, step)
        df.drop(
            ['Event Date', 'Event/Point Description', 'Logical Device', 'Location', 'Channel', 'Panel', 'timestamp'],
            inplace=True, axis=1)
        df.sort_values(['Card Number', 'Date', 'In_Time'], inplace=True, axis=0)
        df_outtime, df_intime = [x for _, x in df.groupby(df['inout'] == " IN")]
        df.drop(['In_Time', 'inout'], axis=1, inplace=True)
        i = df_intime['In_Time']
        o = df_outtime['In_Time']
        df['InTime'] = i
        df['OutTime'] = o
        df['InTime'].fillna(method='ffill', inplace=True)
        df['OutTime'].fillna(method='backfill', inplace=True)
        df.drop_duplicates(keep='first', inplace=True)
        fmt = '%H:%M:%S'
        df['Ins'] = pd.to_datetime(df['InTime'],
                                   format=fmt,
                                   errors='coerce')
        df['Outs'] = pd.to_datetime(df['OutTime'],
                                    format=fmt,
                                    errors='coerce')
        df['OUTSIDE(m)'] = df['Ins'] - df['Outs']
        df['OUTSIDE(m)'] = df['OUTSIDE(m)'] / np.timedelta64(1, 's')
        df['seconds'] = df['OUTSIDE(m)'].abs().round(decimals=1)
        seconds = df['seconds'] % (24 * 3600)
        df['hours'] = df['seconds'] // 3600
        df['seconds'] %= 3600
        df['minutes'] = df['seconds'] // 60
        df['seconds'] %= 60
        df['IN'] = df['hours']+ ':' + df['minutes']+ ':' + df['seconds']
        df['instamp'] = pd.to_datetime(df['IN'], format=fmt,
                                       errors='coerce')
        df['IN'] = df['instamp'].dt.time
        df.drop(['Ins', 'Outs', 'OUTSIDE(m)', 'hours', 'minutes', 'seconds', 'instamp'], inplace=True, axis=1)
        ptf = Frame(home)
        ptf.pack(fill=BOTH, expand=1)
        table = pt = Table(ptf, dataframe=df,
                           showtoolbar=False, showstatusbar=False,editable=False,align='center')
        table.autoResizeColumns()
        table.redraw()
        pt.show()
        


def main():
    global home
    home = Frame(root)
    home.pack(fill=BOTH, expand=1)
    Label(home, text=' WELCOME %s ' % (sysn), font='Arial 16 bold', fg='green').pack(ipady=10, ipadx=10)
    Label(home, text=" EMPLOYEE WORKING HOURS BY SWIPE XLSX/CSV FORMAT DATA ", font='Arial 22 bold',
          fg='darkblue').pack(side='top', ipady=5)
    ub = Button(home, text='UPLOAD FILE', command=lambda: [b_file(), ub.destroy()])
    ub.pack(ipady=5, ipadx=3)
    mb = Menubutton(home, text='Menu', relief='raise', font='arial 12 bold', bg='grey')
    mb.menu = Menu(mb, tearoff=0)
    mb["menu"] = mb.menu
    mb.menu.add_command(label="Home Page", font='arial 10 bold', command=lambda: [home.destroy(), main(), ub.destroy()])
    mb.menu.add_command(label="By Date", font='arial 10 bold', command=by_date)
    mb.menu.add_command(label="To Excel", font='arial 10 bold', command=t_excel)
    mb.menu.add_command(label="To csv", font='arial 10 bold', command=t_csv)
    mb.menu.add_command(label="To Html", font='arial 10 bold', command=t_html)
    mb.pack(padx=15, anchor=NE)
    root.mainloop()


main()
