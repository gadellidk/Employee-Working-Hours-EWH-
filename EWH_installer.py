import os,shutil,winreg,sys,sysconfig,subprocess
from win32com.client import Dispatch
from tkinter import*
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

npath = os.environ['HOMEPATH']+'\EWH'
d = os.environ['HOMEPATH']+'\Desktop\EWH'+'.lnk'
if not os.path.exists(npath):
    root = Tk()
    wwin = 538
    hwin = 338
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    xc = (sw / 2) - (wwin / 2)
    yc = (sh / 2) - (hwin / 2)
    root.overrideredirect(1)
    root.geometry('%dx%d+%d+%d' % (wwin, hwin, xc, yc))
    filename = PhotoImage(file=resource_path('ewh.png') )
    background_label = Label(root, image=filename)
    background_label.place(x=0, y=0, relwidth=1, relheight=1)
    l1 = Label(root,text='Checking ...',font='Times 10 bold',bg = 'skyblue')
    l1.pack(side = 'bottom',fill='x')
    shutil.copytree(resource_path('ewh'),npath)
    l1.after(1000, lambda: l1.destroy())
    subprocess.check_call(['attrib','+H',npath+'/temp.json'])
    packageName = npath
    scriptsDir = npath
    target = os.path.join(scriptsDir, 'EWH.exe')
    linkName =  packageName + '.lnk'
    pathLink = os.path.join(d)
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(pathLink)
    shortcut.Targetpath = target
    l2 = Label(root,text='Installing ...',font='Times 10 bold',bg = 'skyblue')
    l2.pack(side = 'bottom',fill='x')
    l2.after(1500,lambda: l2.destroy())
    shortcut.WorkingDirectory = scriptsDir
    shortcut.IconLocation = target
    shortcut.save()
    root.after(3000, lambda: root.destroy())
    root.mainloop()
else:
    sec_root = Tk()
    wwin, hwin = 500, 100
    sw, sh = sec_root.winfo_screenwidth(),sec_root.winfo_screenheight()
    xc = (sw / 2) - (wwin / 2)
    yc = (sh / 2) - (hwin / 2)
    sec_root.geometry('%dx%d+%d+%d' % (wwin, hwin, xc, yc))
    sec_root.overrideredirect(1)
    sec_root.config(bg='gray25')
    msg = "Uninstalled Succesfully"
    shutil.rmtree(npath)
    os.remove(d)
    Label(sec_root, text=msg, font='Times 20 bold', bg='gray25').pack(padx=2, pady=10)
    Button(sec_root, text='Okay', height=1, width=6, bg='gray25', command=sec_root.destroy).pack(padx=2, pady=4)
    sec_root.mainloop()
    
