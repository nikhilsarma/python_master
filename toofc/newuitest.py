from ofc_copytestv1 import *
#from xl_comp_new import *
from Tkinter import *
import tkFileDialog, random
from tkMessageBox import *
BGCOL = 'aqua'

root = Tk()
root.configure(background=BGCOL)
root.title("The Comparator!")
root.geometry("500x350")
desc = "Welcome to the Comparision tool!"

Label(root, justify=RIGHT, pady = 10, text=desc,bg=BGCOL).grid(row=5)
Label(root, justify=RIGHT, pady = 10,bg=BGCOL, text="Select the directory for hassle free comparision").grid(row=8)
Label(root, text="File Directory: ",bg=BGCOL).grid(row=10)

Label(root, justify=RIGHT, pady = 10, bg=BGCOL,text="OR! select two files to Compare.").grid(row=13)
Label(root, text="Base file : ",bg=BGCOL).grid(row=15)
Label(root, text="Compared to : ",bg=BGCOL).grid(row=17)


e1 = Entry(root)
e2 = Entry(root)
e3 = Entry(root)
e4 = Entry(root)

e1.place(x=200, y=115, width=155)
e2.place(x=200, y=190, width=155)
e3.place(x=200, y=220, width=155)
e4.place(x=145, y=320, width=50)
e4.insert(0, 0)
rep = []

def to_comp():
    
    col = color.get()
    #print comp_mode.get()
    if res_mode.get():
        eps = e4.get()
    else:
        eps = 0

    fd = e1.get()
    f1 = e2.get()
    f2 = e3.get()
    #print col,eps


    if comp_mode.get() == 1:
        res = the_mess1([fd,f1,f2,col,float(eps)])
        #pass
        
    elif comp_mode.get() == 2:
        import fastreader
        res = fastreader.the_lmess([fd,f1,f2])
        #pass
    
    elif comp_mode.get() == 3:
        import test1csv 
        res = test1csv.the_mess([fd,f1,f2])
        print res
    if res:
        print res
        showinfo("Done!","Finished Comparing!")
        rep.append(res)
    

def gen_report():

    #print rep
    
    rtxt_xl = "Exec.time\tWorkBook [Result]\t\tWorkSheets\tResult"
    uxl = "-"*(len(rtxt_xl)+38)
    rtxt_csv = "Exec.time\tfiles [Result]\t\tCreated file\tResult"
    print str('\n\n') + uxl
    
    #msg = "Trust because you are willing to accept the risk, not because it's safe or certain"
    try:
        if comp_mode.get() == 1:
            print rtxt_xl + str('\n')  + uxl
            for e in rep[0]:
                print str(e[0]) + str('\t\t') +  e[1] +str(" Vs ")+e[2]+ " ["+ e[3].upper() + "]"
                rep1(e[4])
        elif comp_mode.get() == 3:
            print rtxt_csv + str('\n')  + uxl
            for e in rep[0]:
                print str(e[0]) + str('\t\t') +  str(e[2])+ " ["+ e[1].upper() + "]"
        if len(rep) != 0:
            del rep[0]
    except Exception as e:
        print "No report generated: " + str(e)
    
def dir_open():
    dir_path = tkFileDialog.askdirectory()
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

res_mode = IntVar()
comp_mode = IntVar()
comp_mode.set(1)

clist = ["Yellow", "Olive", "Aqua","Cocoa","Orange"]
color = StringVar(root)
color.set("Orange")
Button(root, text='select',bg=BGCOL, command=dir_open).grid(row=10, column=3, sticky=NW, pady=4)
Button(root, text='select', bg=BGCOL,command=file_open1).grid(row=15, column=3, sticky=NW, pady=4)
Button(root, text='select', bg=BGCOL, command=file_open2).grid(row=17, column=3, sticky=NW, pady=4)

Button(root, text='compare',bg='brown', command=to_comp).grid(row=20,column=1, sticky=W, pady=4)
Button(root, text='reset', command=to_reset).grid(row=20,column=2, sticky=W, pady=4)
Button(root, text='Gen..report',bg="green", command=gen_report).grid(row=21,column=1, sticky=W, pady=3)
Button(root, text='About',bg="yellow", command='').grid(row=23,column=5, sticky=W, pady=3)
cbb = Checkbutton(root, text="ignore differences <=",bg=BGCOL, variable=res_mode).grid(row=23, sticky=W)
crbC = Radiobutton(root, text="CSV",bg=BGCOL, variable=comp_mode,value=3).place(x=2, y=285,width=40)
crbX = Radiobutton(root, text="XL",bg=BGCOL, variable=comp_mode,value=1).place(x=2, y=260,width=35)
crbVX = Radiobutton(root, text="large XL",bg=BGCOL, variable=comp_mode,value=2).place(x=45, y=260,width=65)

Label(root, text="Mark Color : ",bg=BGCOL).grid(row=7,column=0)

om = OptionMenu(root,color,*clist)
om.config(bg = BGCOL,width=6)
om['menu'].config(bg = BGCOL)
#w.place(x=220, y=30, width=75)
om.grid(row=7,column=1,sticky=W)

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
