#from ofc_copytestv1 import *
#from xl_comp_new import *
#from compxl import *
from Tkinter import *
import ttk
import tkFileDialog, random
from tkMessageBox import *
BGCOL = '#FFDEAD'

root = Tk()
root.configure(background=BGCOL)
root.title("dBase Regrettor!")
root.geometry("500x360")
desc = "Welcome to DataBase Regression tool !"
hline = 85*"-"
Label(root, justify=CENTER, pady = 10, text=desc,bg=BGCOL).place(x=110,y=5)
Label(root, justify=RIGHT, pady = 10,bg=BGCOL, text="Select the file for Regression:").place(x=3,y=60)
Label(root, text=hline,bg=BGCOL).place(x=25, y=170, height = 20)
Label(root, text="Excel File Comparator !",bg=BGCOL).place(x=170, y=185)
Label(root,bg=BGCOL, text="Base File :").place(x=30,y=220)
Label(root,bg=BGCOL, text="Compared to :").place(x=30,y=250)

e1 = Entry(root)
e2 = Entry(root)
e3 = Entry(root)
e4 = Entry(root)

e1.place(x=180, y=69, width=160)

e2.place(x=150, y=220, width=160)
e3.place(x=150, y=250, width=160)
e4.place(x=145, y=325, width=45,height=18)
e4.insert(0, 0)
rep = []

def to_comp():
    f1,f2 = e2.get(),e3.get()
    print f1,f2
    pass
    
def to_reg():
    db1,db2 = env1.get(),env2.get()
    regfile = e1.get()
    if db1 == db2:
        showinfo("Wake up! ","Stop comparing apples - Apples." + "\n" + "     There're Oranges too!")
    print db1,db2,regfile
    
def gen_report():

    rtxt = "Exec.time\tWorkBook [Result]\t\tWorkSheets\tResult"
    u = "-"*(len(rtxt)+38)
    print str('\n\n') + u
    print rtxt + str('\n')  + u
    #msg = "Trust because you are willing to accept the risk, not because it's safe or certain"
    try:
        ecnt = 0
        for e in rep[0]:
            print str(e[0]) + str('\t\t') +  e[1] +str(" Vs ")+e[2]+ " ["+ e[3].upper() + "]"
            rep1(e[4])
        if len(rep) != 0:
            del rep[0]
    except Exception as e:
        print "No report generated: " + str(e)
    
def reg_open():
    dir_path = tkFileDialog.askopenfilename()
    e1.delete(0, END)
    e1.insert(0, dir_path)

def file_open1():
    file_path = tkFileDialog.askopenfilename()
    e2.delete(0, END)
    e2.insert(0, file_path)

def file_open2():
    file_path = tkFileDialog.askopenfilename()
    e3.delete(0, END)
    e3.insert(0, file_path)

def to_reset():
    if len(rep):
        del rep[0]
    e1.delete(0, END)
    e2.delete(0, END)
    e3.delete(0, END)
    res_mode.set(0)
    comp_mode.set(0)
    env1.set("UAT")
    env2.set("PROD")

res_mode = IntVar()
comp_mode = IntVar()


envlist = ["UAT", "PROD", "QA(PreProd)"]
env1,env2 = StringVar(root),StringVar(root)
env1.set("UAT")
env2.set("PROD")
#cbx = StringVar(root)

Button(root, text='select',bg=BGCOL, command=reg_open).place(x=360, y=67,width=50,height=25)
Button(root, text='select',bg=BGCOL, command=file_open1).place(x=340, y=218,width=50,height=24)
Button(root, text='select',bg=BGCOL, command=file_open2).place(x=340, y=248,width=50,height=24)

Button(root, text='Execute ME!',bg='red', command=to_reg).place(x=360, y=125,width=75,height=27)
Button(root, text='Compare',bg='yellow', command=to_comp).place(x=210, y=285,width=75,height=28)
Button(root, text='reset all', command=to_reset).place(x=30, y=285,width=75,height=28)
Button(root, text='Gen..report',bg="green", command=gen_report).place(x=290, y=285,width=75,height=28)
cb_dif = Checkbutton(root, text="ignore differences <=",bg=BGCOL, variable=res_mode).place(x=1, y=323)

Label(root, text="Select Envi.. 1: ",bg=BGCOL).place(x=5, y=100, height = 30, width=110)

en1 = OptionMenu(root,env1,*envlist)
en1.config(bg = BGCOL)
en1['menu'].config(bg = BGCOL)
en1.place(x=15, y=125, height = 25, width=95)


Label(root, text="Select Envi.. 2: ",bg=BGCOL).place(x=120, y=100, height = 30, width=110)

en2 = OptionMenu(root,env2,*envlist)
en2.config(bg = BGCOL)
en2['menu'].config(bg = BGCOL)
en2.place(x=130, y=125, height = 25, width=95)
"""
value = StringVar()
box = ttk.Combobox(root, textvariable=value, state='readonly')
box['values'] = ('A', 'B', 'C')
box.current(0)
box.configure(background= BGCOL,foreground=BGCOL)


#cbox = ttk.Combobox(root,cbx,state='readonly')
#cbox['values']=('a','b','c')
#cbox.current(0)
box.place(x=1, y=2, height = 25, width=96)
"""
mainloop()
