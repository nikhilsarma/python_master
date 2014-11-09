from xl_comp import *
from Tkinter import *
import tkFileDialog

root = Tk()
root.title("Excel Document Bot!")
root.geometry("500x300")
desc = "Welcome to Excel Comparision tool!"

Label(root, justify=RIGHT, pady = 10, text=desc).grid(row=5)
Label(root, justify=RIGHT, pady = 10, text="Select the directory for hastle free comparision").grid(row=8)
Label(root, text="File Directory: ",bg='blue', relief=RIDGE).grid(row=10)

Label(root, justify=RIGHT, pady = 10, text="OR! select two files to Compare.").grid(row=13)
Label(root, text="Base file : ", relief=RIDGE).grid(row=15)
Label(root, text="Compared to : ", relief=RIDGE).grid(row=17)

e1 = Entry(root)
e2 = Entry(root)
e3 = Entry(root)
e1.grid(row=10, column=1)
e2.grid(row=15, column=1)
e3.grid(row=17, column=1)

def to_comp():
    print e1.get()
    the_mess([e1.get(), e2.get(), e3.get()])
    

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

Button(root, text='compare', command=to_comp).grid(row=20,column=1, sticky=W, pady=4)
Button(root, text='select', command=dir_open).grid(row=10, column=3, sticky=W, pady=4)
Button(root, text='select', command=file_open1).grid(row=15, column=3, sticky=W, pady=4)
Button(root, text='select', command=file_open2).grid(row=17, column=3, sticky=W, pady=4)

mainloop()
