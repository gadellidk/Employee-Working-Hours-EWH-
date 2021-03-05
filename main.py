from tkinter import *
def main():
    try:
        import Security_check
    except:
        sec_root = Tk()
        wwin, hwin = 500, 100
        sw, sh = sec_root.winfo_screenwidth(),sec_root.winfo_screenheight()
        xc = (sw / 2) - (wwin / 2)
        yc = (sh / 2) - (hwin / 2)
        sec_root.geometry('%dx%d+%d+%d' % (wwin, hwin, xc, yc))
        sec_root.overrideredirect(1)
        sec_root.config(bg='gray25')
        msg = "Check Something Wrong !"
        Label(sec_root, text=msg, font='Times 20 bold', bg='gray25').pack(padx=2, pady=10)
        Button(sec_root, text='Okay', height=1, width=6, bg='gray25', command=sec_root.destroy).pack(padx=2, pady=4)
        sec_root.mainloop()

main()
