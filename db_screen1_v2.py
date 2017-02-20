"""
__author__ = "Nikhil"
__copyright__ = "LoL"
__maintainer__ = "Team E"
__status__ = "Mess"

"""


from final_v2 import *
from Tkinter import *
import ttk
import tkFileDialog, random
from tkMessageBox import *

"""
Create a root countainer to hold the GUI Objects
The Labels, textboxes, Selector buttons, Dropdown Menus and checkboxes
"""

BGCOL = '#FFDEAD'
root = Tk()
root.configure(background=BGCOL)
root.title("dBase Regrettor!")
root.geometry("500x360")
desc = "Welcome to DataBase Regression tool !"

hline = 85*"-"
Label(root, justify=CENTER, pady = 10, text=desc,bg=BGCOL).place(x=110,y=5)
Label(root, justify=RIGHT, pady = 5,bg=BGCOL, text="Select the file for Regression:").place(x=3,y=50)
Label(root, text=hline,bg=BGCOL).place(x=25, y=170, height = 20)
Label(root, text="Excel File Comparator !",bg=BGCOL).place(x=170, y=185)
Label(root,bg=BGCOL, text="Base File :").place(x=30,y=220)
Label(root,bg=BGCOL, text="Compared to :").place(x=30,y=250)

e1 = Entry(root)
e2 = Entry(root)
e3 = Entry(root)
e4 = Entry(root)
e5 = Entry(root)

e1.place(x=180, y=56, width=160)

e2.place(x=150, y=220, width=160)
e3.place(x=150, y=250, width=160)
e4.place(x=145, y=325, width=45,height=18)
e4.insert(0, 0)
e5.place(x=250, y=123,width=115,height=20)
rep = []

#test function 1
def to_comp():
    f1,f2 = e2.get(),e3.get()
    print f1,f2
    pass

"""
Function to take in the Test
takes in the Testcase number from the textbox and checks them in the
Excel uploaded and runs that particular test case
"""    

def range_list(s):
    l = []
    ast,aed = s[:5],s[6:]
    for i in range(int(ast[2:]),int(aed[2:])+1,1):
        if len(str(i)) == 1:
            l.append('TC00' + str(i))
        elif len(str(i)) == 2:
            l.append('TC0' + str(i))
        elif len(str(i)) == 3:
            l.append('TC' + str(i))
    return l


#The Re-Run function for specific test case execution from the regression file

def re_run():
    
    tlist = e5.get()
    
    if len(tlist) > 5:
        
        if tlist[5] == ':':
            totes = range_list(tlist)
        else:
            totes = tlist.split(',')
            totes.sort()
        #print totes
    else:
        totes = tlist.split(',')
        
    base1,base2 = env1.get(),env2.get()
    regfile = e1.get()

    re_runner(base1,base2,regfile,totes)
    


# Actual funtion starting point after clicking on the "Execute Me" button    

def to_reg():
    
    try:
            
            base1,base2 = env1.get(),env2.get()
            regfile = e1.get()
            if base1 == base2:
                showinfo("Wake up! ","Stop comparing apples - Apples." + "\n" + "     There're Oranges too!")
            print base1,base2,regfile
            runto_excel(base1,base2,regfile)
    except KeyboardInterrupt:
        print "program stopped by user"
    except Exception as e:
        print e


#Generates the report to the Shell in a readable format from a dictonary of values returned as output    
#Only used for Data comparision report alone -Not sure if i have linked these :P
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



"""
Functions to Open the files
reg_open - Opens for folder selection
file_open1/2 - Opens for the base/comp file selections
to_reset - resets the checkboxez/textboxes to default values
"""

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

#Checkboxes are read as 0/1 settings : reset_mode/compare_mode
res_mode = IntVar()
comp_mode = IntVar()



"""
Dropdown list for environment variable selections
The list names depends on the names created in the DSN sources @Desktop
Defauls are set to db_uat Vs. db_prod
"""
envlist = ["db_uat", "db_prod", "db_qa", "QA(PreProd)","db_ui_prod","db_ui_uat"]
env1,env2 = StringVar(root),StringVar(root)
env1.set("db_uat")
env2.set("db_prod")
#cbx = StringVar(root)

#Buttons & placement for select options
Button(root, text='select',bg=BGCOL, command=reg_open).place(x=360, y=54,width=50,height=25)
Button(root, text='select',bg=BGCOL, command=file_open1).place(x=340, y=218,width=50,height=24)
Button(root, text='select',bg=BGCOL, command=file_open2).place(x=340, y=248,width=50,height=24)

#Buttons & placement for special Functions
Button(root, text='Execute ME!',bg='red', command=to_reg).place(x=400, y=90,width=75,height=27)
Button(root, text='Re Run!',bg='red', command=re_run).place(x=390, y=135,width=55,height=20)
Button(root, text='Compare',bg='yellow', command=to_comp).place(x=210, y=285,width=75,height=28)
Button(root, text='reset all', command=to_reset).place(x=30, y=285,width=75,height=28)
Button(root, text='Gen..report',bg="green", command=gen_report).place(x=290, y=285,width=75,height=28)

#Ignore differences check-box
cb_dif = Checkbutton(root, text="ignore differences <=",bg=BGCOL, variable=res_mode).place(x=1, y=323)

#Dropdown menu for Environment-1 selection
Label(root, text="Select Envi.. 1: ",bg=BGCOL).place(x=5, y=100, height = 30, width=110)
en1 = OptionMenu(root,env1,*envlist)
en1.config(bg = BGCOL)
en1['menu'].config(bg = BGCOL)
en1.place(x=15, y=125, height = 25, width=95)

#Dropdown menu for Environment-2 selection
Label(root, text="Select Envi.. 2: ",bg=BGCOL).place(x=120, y=100, height = 30, width=110)
en2 = OptionMenu(root,env2,*envlist)
en2.config(bg = BGCOL)
en2['menu'].config(bg = BGCOL)
en2.place(x=130, y=125, height = 25, width=95)

Label(root, text= "Test Case No: ",bg=BGCOL).place(x=230, y=100)

#start the mainloop of the Tinkter container
mainloop()
