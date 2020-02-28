from tkinter import *
from tkinter import filedialog
import pandas as pd
from dictionary import *
from extracter import *

window = Tk()
window.title('RTI WAWA Inventory')
window.geometry('300x200')

product_num = []
asset_tag = []
serial = []

def find_file():
    global df
    file = filedialog.askopenfilename(filetypes =[("Excel files", ".xlsx .xls .xlsm")])
    df = pd.read_excel(file,skiprows=14)
    print('file done')
    return df

def run_program():
    if len(scan1.get()) != 0:
        find_pn()
        print('Start Check')
        pn_check(df)
        print('Finished')
        append_file()

    else:
        print('start_check fail')

def find_pn():
    global product_num
    for item in rti_remodel_product.items():
        if item[0] == scan1.get():
            product_num = item[1]
            print('found')
            return product_num

def pn_check(df):
    for n in df['Product ID[*]']:
        if n == product_num:
            print(product_num)
            x = df.loc[df['Product ID[*]'] == product_num, ['Tag No','Serial No.[*]']]
            if str(x['Tag No'].iloc[0])=='nan':
                asset_tag = scan2.get()
                serial = scan3.get()
                x['Tag No'].iloc[0] = asset_tag
                x['Serial No.[*]'] = serial
                print(serial)
                break
            else:
                continue

def append_file():
    exec(open("extracter.py").read())


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