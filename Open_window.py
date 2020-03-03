from tkinter import *
from tkinter import filedialog
from dictionary import *
from openpyxl import *

product_num = []

def find_file():
    global workbook
    global wb
    global file_name
    file = filedialog.askopenfilename(filetypes =[("Excel files", ".xlsx .xls .xlsm")])
    file_name = file.split('/')[-1]
    wb = load_workbook(file,data_only=True,keep_vba=True)
    workbook = wb.active
    print('file done')
    return workbook

def run_program():
    if len(scan1.get()) != 0:
        find_pn()
        print('Start Check')
        pn_check(workbook)
        print('Finished')
        print(file_name)
        wb.save(filename='New_'+ file_name)
        print('Saved')

def find_pn():
    global product_num
    print(scan1.get())
    for item in rti_remodel_product.items():
        if item[0] == scan1.get():
            product_num = item[1]
            print('found')
            print(product_num)
            return product_num

def pn_check(workbook):
    for row in workbook.iter_rows():
        for cell in row:
            if cell.value == product_num:
                print(str(cell.row))
                if workbook[cell.row][3].value is None and workbook[cell.row][4].value is None:
                    workbook[cell.row][3].value = scan2.get()
                    workbook[cell.row][4].value = scan3.get()
                    print(workbook[cell.row][3].value)
                    print('Check finished ')
                    break
                else:
                    continue

top = Toplevel(window)

top_btn = Button(top, text='Top button',command=find_file)

window = Tk()
window.title('RTI WAWA Inventory')
window.geometry('300x200')

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


top.mainloop()
window.mainloop()