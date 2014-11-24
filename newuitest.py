from ofc_copy import *
#from xl_comp_new import *
from Tkinter import *
import tkFileDialog, random
from tkMessageBox import *
BGCOL = 'aqua'

root = Tk()
root.configure(background=BGCOL)
root.title("Excel Comparator!")
root.geometry("500x350")
desc = "Welcome to Excel Comparision tool!"

Label(root, justify=RIGHT, pady = 10, text=desc,bg=BGCOL).grid(row=5)
Label(root, justify=RIGHT, pady = 10,bg=BGCOL, text="Select the directory for hassle free comparision").grid(row=8)
Label(root, text="File Directory: ",bg=BGCOL, relief=RIDGE).grid(row=10)

Label(root, justify=RIGHT, pady = 10, bg=BGCOL,text="OR! select two files to Compare.").grid(row=13)
Label(root, text="Base file : ",bg=BGCOL, relief=RIDGE).grid(row=15)
Label(root, text="Compared to : ",bg=BGCOL, relief=RIDGE).grid(row=17)


e1 = Entry(root)
e2 = Entry(root)
e3 = Entry(root)
e1.place(x=220, y=115, width=145)
e2.place(x=220, y=190, width=145)
e3.place(x=220, y=220, width=145)

rep = []
def to_comp():
    
    col = color.get()
    print col
    fd = e1.get()
    f1 = e2.get()
    f2 = e3.get()
    try:
        
        if f1.endswith('.xlsx') != f2.endswith('.xlsx'):
            showwarning("Something happend! ","Cannot Compare different files")             
        if fd == '' and f1 == '' and f2 == '':
            showerror("Damn! ", "Please select files to compare")
        elif fd == '' and (f1 == '' or f2 == ''):
            showwarning("Uh Oh!", "Enter either a folder OR The two files to compare")
        else:
            res = the_mess([fd,f1,f2,col])
            if res:
                showinfo("Done!","Finished Comparing!")
            rep.append(res)
    except Exception as exp:
        showwarning("Something happend! ",str(exp.__class__.__name__) + ": " +str(exp.message))
def gen_report():

    print rep
    
    rtxt = "Exec.time\tWorkBook [Result]\t\tWorkSheets\tResult"
    u = "-"*(len(rtxt)+32)
    print str('\n\n') + u
    print rtxt + str('\n')  + u
    #msg = "Trust because you are willing to accept the risk, not because it's safe or certain"
    try:
        ecnt = 0
        for e in rep[0]:
            print str(e[0]) + str('\t\t') +  e[1] +str(" Vs ")+e[2]+ " ["+ e[3].upper() + "]"
            rep1(e[4])
            #print e[3],
            if e[3] == "Fail":
                ecnt += 1
        if ecnt >= 1:
            #msg = random.choice(freport)
            print str('\n') + "Oops!" + str('\n') + random.choice(freport)
        else:
            print str('\n') + "Cool! but," + str('\n') + msg
        if len(rep) != 0:
            del rep[0]
    except Exception as e:
        print "No report generated" + str(e)
    
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

#res_mode = IntVar()
clist = ["Yellow", "Olive", "Aqua","Cocoa"]
color = StringVar(root)
color.set("Yellow")
Button(root, text='select',bg=BGCOL, command=dir_open).grid(row=10, column=3, sticky=NW, pady=4)
Button(root, text='select', bg=BGCOL,command=file_open1).grid(row=15, column=3, sticky=NW, pady=4)
Button(root, text='select', bg=BGCOL, command=file_open2).grid(row=17, column=3, sticky=NW, pady=4)

Button(root, text='compare',bg='brown', command=to_comp).grid(row=20,column=1, sticky=W, pady=4)
Button(root, text='reset', command=to_reset).grid(row=20,column=2, sticky=W, pady=4)
Button(root, text='Gen..report',bg="green", command=gen_report).grid(row=21,column=1, sticky=W, pady=3)
Button(root, text='About',bg="yellow", command='').grid(row=23,column=5, sticky=W, pady=3)
#Checkbutton(root, text="write diff. to new file", variable=res_mode).grid(row=23, sticky=W)
Label(root, text="Mark Color : ",bg=BGCOL, relief=RIDGE).grid(row=7,column=0)
OptionMenu(root,color,*clist).grid(row=7,column=1,sticky=W)
#OptionMenu.configure(bg='green')
#cw["menu"].config(bg="GREEN")
#Report generation
def rep1(dic):
    for k,v in dic.items():
        print str(k).rjust(60) + str('\t') + str(v[0]) + str(': ')+str(v[1])


def rep3(dic):
	for k,v in dic.items():
		print str("\t\t\t\t\t") + str(k) + str("\t\t") + str(v[0]) + str("\t") + str(v[1])

#report with errors and exceptions(shows the exception name)
def rep2(dic):
	for k,v in dic.items():
		print str("\t\t\t\t\t") + str(k) + str("\t\t") + str(v[0]) + str("\t") + str(v[1].__class__.__name__) + str(": ") + str(v[1])

freport = ["I'm not afraid of the cemetery. It's the only place the ghosts don't follow me...", 
"The lights flicker. I put the pillow over my head, so I won't hear the screams this time...",
"I burned the dolls even though my children cried. They did not understand my fear because they assumed I was who moved the dolls into their beds each night...",
"She asked why I was breathing so heavily. I wasn't...",
'''"I can't sleep" she whispered, crawling into bed with me. I woke up cold, clutching the dress she was buried in...''']


mainloop()
