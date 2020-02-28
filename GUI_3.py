from tkinter import *
from tkinter import filedialog
import pandas as pd
from dictionary import *
from extracter import *
from openpyxl import *

window = Tk()
window.title('RTI WAWA Inventory')
window.geometry('300x200')

product_num = []

def find_file():
    global workbook
    file = filedialog.askopenfilename(filetypes =[("Excel files", ".xlsx .xls .xlsm")])
    wb = load_workbook(file,data_only=True)
    workbook = wb.active
    print('file done')
    return workbook

def run_program():
    if len(scan1.get()) != 0:
        find_pn()
        print('Start Check')
        pn_check(df)
        print('Finished')
        wb.save(file)

def find_pn():
    global product_num
    for item in rti_remodel_product.items():
        if item[0] == scan1.get():
            product_num = item[1]
            print('found')
            return product_num

def pn_check(workbook):
    for value in workbook.values:
        if value[0] == 'RWS':
            continue
        if value[6] == product_num:
            print(value[6])
            x = value[3:5]
            print(x[0])
            if x[0] is None and x[1] is None:
                location = value[0]
                sheet[location][3].value = scan2.get()
                sheet[location][4].value = scan3.get()
                print(sheet[location][4].value)
                print('Check_finished')
                break


            else:
                continue

#Add file button
btn = Button(window,text='Add File', command = find_file)
btn.grid(column=5,row=1)

#Text to indicate file browsing
lbl1 = Label(window,text='Product Number')
lbl1.grid(column=4,row=2)

#Entry field for product number
scan1 = Entry(window, width=15)
scan1.grid(column=5,row=2)

#product number submit button
btn1 = Button(window,text='Add',command =run_program)
btn1.grid(column=5,row=5)

#Entry field for asset tags
scan2 = Entry(window, width=15)
scan2.grid(column=5,row=3)
#Text to indicate asset tag field
lbl2 = Label(window,text='Asset Tag')
lbl2.grid(column=4,row=3)

#Entry field for serial number
scan3 = Entry(window, width=15)
scan3.grid(column=5,row=4)
#Text to indicate Serial Number Field
lbl3 = Label(window,text='Serial Number')
lbl3.grid(column=4,row=4)





window.mainloop()