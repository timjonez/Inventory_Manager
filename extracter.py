
import pandas as pd
from zipfile import ZipFile
from GUI_2 import  df


writer = pd.ExcelWriter('test.xlsx', engine='xlsxwriter')

df.to_excel(writer)
vba_filename = 'vbaProject.bin'
xlsm_file = 'Remodel.xlsm'

xlsm_zip = ZipFile(xlsm_file, 'r')

# Read the xl/vbaProject.bin file.
vba_data = xlsm_zip.read('xl/' + vba_filename)

# Write the vba data to a local file.
vba_file = open(vba_filename, "wb")
vba_file.write(vba_data)
vba_file.close()

workbook  = writer.book
workbook.filename = 'test.xlsm'
workbook.add_vba_project('./vbaProject.bin')

writer.save()