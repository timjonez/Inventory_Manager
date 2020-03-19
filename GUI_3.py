from tkinter import filedialog, messagebox, Listbox
import tkinter as tk
from dictionary import *
from openpyxl import *
import atexit

window = tk.Tk()
window.title('WAWA Scanner')
window.geometry('300x250')

picked = tk.StringVar()
picked.set('----')
product_num = []

def find_file():
    global workbook
    global wb
    global file_name
    global file_path
    file = filedialog.askopenfilename(filetypes =[("Excel files", ".xlsx .xls .xlsm")])
    file_name = file.split('/')[-1]
    file_path = file[0:-len(file_name)]
    print(file_path)
    wb = load_workbook(file,data_only=True,keep_vba=True)
    workbook = wb.active
    print('file done')
    if file_name == '':
        print('Try Again')
    else:
        # Show file selected
        show_file = tk.Label(window, text=file_name[0:30]+'...')
        show_file.grid(column=4,row=2,columnspan=3,sticky= tk.W+tk.E)
        if user_select.get() == 1:
            options = tk.OptionMenu(window, picked, *choices)
            options.grid(column=6, row=3)


    scan1.focus_set()
    return workbook

def run_program_remodel():
    if len(scan1.get()) != 0:
        find_pn_remodel()
        print('Start Check')
        pn_check(workbook)
        print('Finished')

def run_program_new():
    global product_num
    if len(scan1.get()) != 0:
        for item in rti_new_product.items():
            if item[0] == scan1.get():
                product_num = item[1]
                print(product_num)
                pn_check(workbook)




def find_pn_remodel():
    global product_num
    print(scan1.get())
    for item in rti_remodel_product.items():
        if item[0] == scan1.get():
            product_num = item[1]
            print('found')
            print(product_num)
            return product_num

def pn_check(workbook):
    if len(product_num)==0:
        tk.messagebox.showinfo("Can't Find","I Can't find this product number. Double-check that it is correct")

    else:
        for row in workbook.iter_rows():
            if row[0].value in exclude:
                print('206/9 Exclude')
                continue
            for cell in row:
                if cell.value == product_num:
                    print(str(cell.row))
                    if workbook[cell.row][3].value is None and workbook[cell.row][4].value is None:
                        ok = tk.messagebox.askokcancel("Item Added", "Add Item as: "+ str(workbook[cell.row][0].value))
                        if not ok:
                            continue
                        workbook[cell.row][3].value = scan2.get()
                        workbook[cell.row][4].value = scan3.get()
                        print(workbook[cell.row][3].value)
                        print('Check finished ')
                        break
            else:
                continue
            break
        else:
            print('No More Nones')
            tk.messagebox.showinfo("Can't Add", "I can't find a empty slot for this item. ")

user_select = tk.IntVar()
user_select.set(0)
def selection():
    if user_select.get() ==1:
        print('New Store')
        if scan1.get() == '':
            pn_lookup()
        run_program_new()
    elif user_select.get() ==2:
        print('Remodel')
        run_program_remodel()
    else:
        print('Failed at selection')

def pn_lookup():
    for item in choices:
        if item == picked.get():
            scan1.insert(0,item)
            print('PN_Found')

def continue_selection_same():
    scan2.delete(first=0,last=100)
    scan3.delete(first=0,last=100)
    print('cleared')
    scan2.focus_set()

def continue_selection_diff():
    scan1.delete(first=0,last=100)
    scan2.delete(first=0,last=100)
    scan3.delete(first=0,last=100)
    print('cleared')
    scan1.focus_set()

def finish_program():
    print(file_name)
    wb.save(filename= file_path+'Filled_' + file_name)
    print('Saved')
    window.destroy()

def exit_prompt():
    comfirm =messagebox.askyesno('Save File?','Do you want to save the file?')
    if comfirm:
        finish_program()
    else:
        print('Closed')

#Add selection for Remodel
Rbtn = tk.Radiobutton(window,text='New Store',indicatoron=0 , variable=user_select, value=1,width=10)
Rbtn.grid(column=4,row=0)

#Add selection for New
Rbtn = tk.Radiobutton(window,text='Remodel',indicatoron=0, variable=user_select, value=2,width=10)
Rbtn.grid(column=6,row=0)

# Add file button
btn = tk.Button(window, text='Add File', command=find_file)
btn.grid(column=5, row=1)

#Text to indicate file browsing
lbl1 = tk.Label(window,text='Product Number')
lbl1.grid(column=4,row=3)

#Entry field for product number
scan1 = tk.Entry(window, width=15)
scan1.grid(column=5,row=3)

#Entry field for asset tags
scan2 = tk.Entry(window, width=15)
scan2.grid(column=5,row=4)
#Text to indicate asset tag field
lbl2 = tk.Label(window,text='Asset Tag')
lbl2.grid(column=4,row=4)

#Entry field for serial number
scan3 = tk.Entry(window, width=15)
scan3.grid(column=5,row=5)
#Text to indicate Serial Number Field
lbl3 = tk.Label(window,text='Serial Number')
lbl3.grid(column=4,row=5)

#product number submit button
btn1 = tk.Button(window,text='Add',command =selection,padx=40)
btn1.grid(column=5,row=6,pady=5)

space = tk.Label(window,text='                 ')
space.grid(column=5,row=7)

#Button to continue adding additional of the same
cont_btn_same = tk.Button(window,text='Add Additional',command= continue_selection_same)
cont_btn_same.grid(column=4,row=8)

#Button to continue adding of other products
cont_btn_diff = tk.Button(window,text='Add Other',command= continue_selection_diff)
cont_btn_diff.grid(column=5,row=8)

#Button to quit adding and finish
fin_btn = tk.Button(window,text='Finish',command=finish_program)
fin_btn.grid(column=6,row=8)


#atexit.register(exit_prompt)


window.mainloop()