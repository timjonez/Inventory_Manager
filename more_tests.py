from tkinter import messagebox
from tkinter import *
from tkinter.ttk import Progressbar
from tkinter import ttk
from tkinter import filedialog

window = Tk()
window.title('Title')
window.geometry('200x200')

selected = IntVar()

def clicked():
    #messagebox.showinfo('This is a Title','This is the fine print')
    #messagebox.showwarning('Title','ERROR!!!!!')
    #messagebox.askretrycancel('Title','Message')
    #messagebox.askyesno('Title', 'Message')
    res = messagebox.askyesnocancel('Title', 'Message')

#file = filedialog.askopenfilename(filetypes =['Excel', ('*.xls', '*.xlsx')])
file = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
menu = Menu(window)
new_item = Menu(menu)
new_item.add_command(label='New')
menu.add_cascade(label='File',menu=new_item)
window.config(menu=menu)

btn = Button(window, text='Click',command=clicked)

btn.grid(column=2,row=1)

var = IntVar()
var.set(5)
spin = Spinbox(window, from_=0, to_=10,width=5,textvariable=var)
spin.grid(column=4,row=1)
"""
style = ttk.Style()
style.theme_use('default')
style.configure("black.Horizonal.TProgressbar", background='blue')
bar = Progressbar(window,length=200, style='black.Horizontal.TProgressbar')
bar['value'] = 70
bar.grid(column=1,row=9)
"""
window.mainloop()
