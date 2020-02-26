
from tkinter import *
from tkinter.ttk import *
from tkinter import scrolledtext

window = Tk()
window.title('RTI')
window.geometry('350x200')

lbl = Label(window, text='Hi RTI')

lbl.grid(column=1, row=1)

txt = Entry(window,width=15)

txt.grid(column=1,row=2)

def clicked():
    res = 'Welcome ' + txt.get()
    lbl.configure(text=res)

btn = Button(window,text='Add name:',command=clicked)
btn.grid(column=1,row=3)

combo = Combobox(window)
combo['values']=('Item1','Item2','Item3')
combo.current(0)
combo.grid(column=3, row=2)

check_state = BooleanVar()
check_state.set(True)
check = Checkbutton(window,text='Yes?', var=check_state)
check.grid(column=1, row=4)

selected = IntVar()

radio1 = Radiobutton(window,text='First', value=1, variable=selected)
radio2 = Radiobutton(window,text='Second',value=2,variable=selected)
radio1.grid(column=3,row=1)
radio2.grid(column=4,row=1)

def click():
    print(selected.get())

rad_btn = Button(window,text='Submit',command=click)
rad_btn.grid(column=4,row=0)

txt = scrolledtext.ScrolledText(window,width=40,height=10)
txt.grid(column=3,row=5)
txt.insert(INSERT,'This is where text will go')









window.mainloop()