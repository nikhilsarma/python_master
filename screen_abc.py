"""
__author__ = "Nikhil Kumar Kadayinti"
__copyright__ = "LoL"
__maintainer__ = "Team E"
__status__ = "Mess"


A simple UI using Tkinter for the Data Comp Program 
"""


from compxl_abc import *
from Tkinter import *
import tkFileDialog, random
from tkMessageBox import *
from tkSimpleDialog import *

"""
Create a root countainer to hold the GUI Objects
The Labels, textboxes, Selector buttons, Dropdown Menus and checkboxes
"""
BGCOL = '#FFDEAD'
root = Tk()
root.configure(background=BGCOL)
root.title("Excel Comparator!")
root.geometry("500x350")
desc = "Welcome to Excel Comparision tool !!"

Label(root, justify=RIGHT, pady = 10, text=desc,bg=BGCOL).place(x=145, y=2)
Label(root, justify=RIGHT, pady = 10, bg=BGCOL, text="Select the directory for hassle free comparision...").place(x=10, y=35)
Label(root, text="File Directory : ",bg=BGCOL).place(x=8, y=78, width=170)

Label(root, justify=RIGHT, bg=BGCOL,text="OR! select two files to Compare...").place(x=10, y=115)
Label(root, text="Base file : ",bg=BGCOL).place(x=8, y=145, width=170)
Label(root, text="Compared to : ",bg=BGCOL).place(x=8, y=172, width=170)


e1 = Entry(root)
e2 = Entry(root)
e3 = Entry(root)
e4 = Entry(root)
e5 = Entry(root)
e6 = Entry(root)
e7 = Entry(root)

e1.place(x=170, y=80, width=170)
e2.place(x=170, y=145, width=170)
e3.place(x=170, y=175, width=170)

#e4.insert(0, 0)
rep = []

def to_comp():
    
    #aaaaa = tkSimpleDialog.askstring("User Creds", "Enter your name", parent=root)

    val1,val2 = e5.get(),e6.get()
    clr1,clr2,clr3 = c1.get(),c2.get(),c3.get()
    cmpmode = comp_mode.get()
    print clr1,clr2,clr3
    print val1,val2
    #pass
    if ignore_mode.get():
        eps_per = e4.get()
        eps_val = e7.get()
        #print eps
    else:
        eps_per = 0
        eps_val = 0
    #print eps_per, eps_val
    
    fd = e1.get()
    f1 = e2.get()
    f2 = e3.get()
    import compxl_abc
    res = compxl_abc.the_mess1([fd,f1,f2,clr1,clr2,clr3,val1,val2,float(eps_per),float(eps_val)])
    print res
    if comp_mode.get() == 1:
        import compxl
        res = compxl.the_mess1([fd,f1,f2,clr1,clr2,clr3,val1,val2,float(eps_per),float(eps_val)])
        #res = the_mess1([fd,f1,f2,col,float(eps_per),float(eps_val)])
        print res
        #pass
        
    elif comp_mode.get() == 2:
        import fastreader
        res = fastreader.the_lmess([fd,f1,f2,col,float(eps_per),float(eps_val)])
        pass
    
    elif comp_mode.get() == 3:
        import cmpcsv
        res = cmpcsv.comp_csv(f1,f2)
        #res = cmpcsv.comp_csv(f1,f2)
        #print res
        
    if res:
        showinfo("Done!","Finished Comparing!")
        rep.append(res)


#Generates the report to the Shell in a readable format from a dictonary of values returned as output    

def gen_report():

    #print rep
    
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

def dir_open():
    dir_path = tkFileDialog.askdirectory()
    e1.delete(0, END)
    e1.insert(0, dir_path)

def file_open1():
    file_path = tkFileDialog.askopenfilename()
    file_name = file_path[file_path.rfind("/")+1:]
    e2.delete(0, END)
    e2.insert(0, file_path)

def file_open2():
    file_path = tkFileDialog.askopenfilename()
    file_name = file_path[file_path.rfind("/")+1:]
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

#Checkboxes are read as 0/1 settings : reset_mode/compare_mode
ignore_mode = IntVar()
comp_mode = IntVar()

"""
Dropdown list for Color selections
Defauls are set to Yellow, Blue, Rose
"""
clist = ["Yellow", "Rose", "Blue","Green","Orange"]
c1,c2,c3 = StringVar(root),StringVar(root),StringVar(root)
c1.set("Yellow")
c2.set("Blue")
c3.set("Rose")
#color1.set("Orange")

#Buttons & placement for select options
Button(root, text='select',bg=BGCOL, command=dir_open).place(x=360, y=77, width=70, height = 25)
Button(root, text='select', bg=BGCOL,command=file_open1).place(x=360, y=143, width=70, height = 25)
Button(root, text='select', bg=BGCOL, command=file_open2).place(x=360, y=173, width=70, height = 25)

#Buttons & placement for special Functions
Button(root, text='compare',bg='brown', command=to_comp).place(x=200, y=215, width=75, height = 28)
Button(root, text='reset', command=to_reset).place(x=385, y=215, width=50, height = 28)
Button(root, text='Gen..report',bg="green", command=gen_report).place(x=295, y=215, width=80, height = 28)
#Button(root, text='About',bg="yellow", command='').grid(row=23,column=5, sticky=W, pady=3)
cbb = Checkbutton(root, text="Ignore    dif_per <=",bg=BGCOL, variable=ignore_mode).place(x=5, y=250)
e4.place(x=135, y=255, width=40, height = 18)
Label(root, text=" Or   dif_val <=",bg=BGCOL).place(x=175, y=253)
e7.place(x=257, y=255, width=40, height = 18)

#crbC = Radiobutton(root, text="CSV",bg=BGCOL, variable=comp_mode,value=3).place(x=2, y=285,width=40)
#crbX = Radiobutton(root, text="XL",bg=BGCOL, variable=comp_mode,value=1).place(x=2, y=260,width=35)
#crbVX = Radiobutton(root, text="large XL",bg=BGCOL, variable=comp_mode,value=2).place(x=45, y=260,width=65)

#e5.place(x=2, y=260, width=50)

#Stuff for 1st Difference selections
Label(root, text="diff's <= : ",bg=BGCOL).place(x=5, y=280, height = 28, width=55)
e5.place(x=60, y=284, width=50, height = 18)
col1 = OptionMenu(root,c1,*clist)
col1.config(bg = BGCOL)
col1['menu'].config(bg = BGCOL)
col1.place(x=15, y=305, height = 25, width=80)

#Stuff for 2nd Difference selections
Label(root, text="diff's <= : ",bg=BGCOL).place(x=130, y=280, height = 28, width=55)
e6.place(x=185, y=284, width=50, height = 18)
col2 = OptionMenu(root,c2,*clist)
col2.config(bg = BGCOL)
col2['menu'].config(bg = BGCOL)
col2.place(x=140, y=305, height = 25, width=80)

#Stuff for 3rd Difference selections
Label(root, text="Rest diff's : ",bg=BGCOL).place(x=270, y=280, height = 28)
col3 = OptionMenu(root,c3,*clist)
col3.config(bg = BGCOL)
col3['menu'].config(bg = BGCOL)
col3.place(x=265, y=305, height = 25, width=80)

#placeholders
e4.insert(0,0)#diff_percentage
e7.insert(0,1)#diff_value
e5.insert(0,1)#1st diff interval
e6.insert(0,10)#2nd diff interval


#Test Fucntions_for research of different result formats
#Label(root, text="Mark Color : ",bg=BGCOL).grid(row=7,column=0)

#om = OptionMenu(root,color,*clist)
#om.config(bg = BGCOL,width=6)
#om['menu'].config(bg = BGCOL)
#w.place(x=220, y=30, width=75)
#om.grid(row=7,column=1,sticky=W)

#Report generation
def rep1(dic):
    for k,v in dic.items():
        print str(k).rjust(60) + str('\t') + str(v[0]) + str(': ')+str(v[1])
    #print str('\n\n') + u.gen_report

def rep3(dic):
	for k,v in dic.items():
		print str("\t\t\t\t\t") + str(k) + str("\t\t") + str(v[0]) + str("\t") + str(v[1])

#report with errors and exceptions(shows the exception name)
def rep2(dic):
	for k,v in dic.items():
		print str("\t\t\t\t\t") + str(k) + str("\t\t") + str(v[0]) + str("\t") + str(v[1].__class__.__name__) + str(": ") + str(v[1])



mainloop()
