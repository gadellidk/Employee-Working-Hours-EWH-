from tkinter import *
from tkinter import messagebox
import pandas as pd
import os,socket
from tkinter.filedialog import askopenfilename
from pandastable import Table
import time
import tempfile

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
class My_Gui_B:
    def paths(self):
        global d_path,f_path
        dt= time.strftime('%d-%m-%y')
        self.t_path = tempfile.gettempdir()
        d_path = os.environ['HOMEPATH']+'\Desktop\EWH_Output_Files'
        f_path = os.environ['HOMEPATH']+'\Desktop\EWH_Output_Files\%s/'%(dt)
        if not os.path.exists(d_path):
            os.makedirs(d_path)
        if not os.path.exists(f_path):
            os.makedirs(f_path)
    def t_html(self):
        if self.br_file is not None:
            self.df.to_html(f_path+"working_hours.html")
            self.e.to_html(f_path+'Working_Hours_By_Date.html')
            messagebox.showinfo("Html" ,"Sucessfully coverted to HTML Files ")
    def t_excel(self):
        if self.br_file is not None:
            self.df.to_excel(f_path+"working_hours.xlsx")
            self.e.to_excel(f_path+'Working_Hours_By_Date.xlsx')
            messagebox.showinfo("Excel"," Sucessfully coverted to Excel Files ")
    def t_csv(self):
        if self.br_file is not None:
            self.df.to_csv(f_path+"working_hours.csv")
            self.e.to_csv(f_path+'Working_Hours_By_Date.csv')
            messagebox.showinfo("csv"," Sucessfully coverted to CSV Files ")
    def by_date(self):
        self.ptf.destroy()
        self.ptb = Frame(self.window)
        self.ptb.pack(fill=BOTH, expand=1)
        table = pt = Table(self.ptb, dataframe=self.e,
                           showtoolbar=False, showstatusbar=False,editable=False,align='center')
        table.autoResizeColumns()
        table.redraw()
        pt.show()
    def b_file(self):
        self.br_file = askopenfilename()
        self.data_frame()
    def data_frame(self):
        if self.br_file is not None:
            try:
                df = pd.read_csv("%s" % (self.br_file),usecols=['Event Date','Location','Card Number'],dtype=object)
            except:
                messagebox.showerror("Format Error !","Please provide csv Format Data File")
        #df = pd.read_csv(r"C:\Users\gdkvarma\Documents\My Projects\Time_Calculations-master\Timedata.csv",usecols=['Event Date','Location','Card Number'],dtype=object)
        df['Event Date'] = pd.to_datetime(df['Event Date'])
        df['IN/OUT']= [io.split(" ")[-1] for io in df['Location']]
        df.sort_values(['Card Number', 'Event Date'], inplace=True, axis=0)
        df.drop(['Location'],inplace=True, axis=1)
        df['Date'] = df['Event Date'].dt.date
        df_outtime, df_intime = [x for _, x in df.groupby(df['IN/OUT'] == "IN")]
        i,o = df_intime['Event Date'],df_outtime['Event Date']
        df['In'],df['Out'] = i,o
        df.drop(['IN/OUT','Event Date'],inplace=True, axis=1)
        df['In'].fillna(method='ffill', inplace=True)
        df['Out'].fillna(method='backfill', inplace=True)
        df.drop_duplicates(keep='first', inplace=True)
        df['Working Hours'] = df['Out'] - df ['In']
        df['In'],df['Out'] = df['In'].dt.time,df['Out'].dt.time
        d = df.groupby(["Card Number","Date"],as_index=False).sum()
        d['Working Hours'] = [w.split(" ")[-1] for w in d['Working Hours'].astype(str)]
        d['Date'] = [w.split(" ")[0] for w in d['Date'].astype(str)]
        e = d.pivot('Card Number', 'Date').rename_axis().fillna('A')
        e.to_csv(self.t_path+"/by_date.csv")
        self.e = pd.read_csv(self.t_path+"/by_date.csv")
        df['Working Hours'] = [w.split(" ")[-1] for w in df['Working Hours'].astype(str)]
        self.df = df
        self.ptf = Frame(self.window)
        self.ptf.pack(fill=BOTH, expand=1)
        table = pt = Table(self.ptf, dataframe=df,
                           showtoolbar=False, showstatusbar=False,editable=False,align='center')
        table.autoResizeColumns()
        table.redraw()
        pt.show()
        
class My_Gui(My_Gui_B):
    def __init__(self):
        self.window = Tk()
        self.width  = self.window.winfo_screenwidth()
        self.height = self.window.winfo_screenheight()
        self.window.geometry('%sx%s+%s+%s' % (self.width, self.height, 0, 0))
    def open_file(self):
        sysn = socket.gethostname()
        Label(self.window, text=' WELCOME %s ' % (sysn), font='NONE 16 bold', fg='green').pack(ipady=10, ipadx=10)
        Label(self.window, text=" EMPLOYEE WORKING HOURS BY SWIPE DETAILS (.CSV)", font='NONE 22 bold',
          fg='darkblue').pack(side='top', ipady=5)
        ub = Button(self.window, text='UPLOAD FILE', command=self.b_file)
        ub.pack(ipady=5, ipadx=3)
        mb = Menubutton(self.window, text='Menu', relief='raise', font='arial 12 bold', bg='grey')
        mb.menu = Menu(mb, tearoff=0)
        mb["menu"] = mb.menu
        mb.menu.add_command(label="Home Page", font='none 10 bold', command=lambda: [self.ptf.destroy(),ub.destroy(),self.run])
        mb.menu.add_command(label="By Date", font='none 10 bold', command=lambda:[self.by_date()])
        mb.menu.add_command(label="To Excel", font='none 10 bold', command=self.t_excel)
        mb.menu.add_command(label="To csv", font='none 10 bold', command=self.t_csv)
        mb.menu.add_command(label="To Html", font='none 10 bold', command=self.t_html)
        mb.pack(padx=15, anchor=NE)
    def run(self):
        self.window.mainloop()
        
        
cl = My_Gui()
cl.open_file()
cl.paths()
cl.run()
