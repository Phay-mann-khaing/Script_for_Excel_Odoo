import csv2xlsx
import os
from Tkinter import Tk
from tkFileDialog import askopenfilename
import convert_into_importable

Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
ft = [('Excel file', '*.xlsx'), ('All Files', '*')]
filename = askopenfilename(filetypes=ft)  # show an "Open" dialog box and return the path to the selected file
# print(filename)

# filename = raw_input("File Name: ")
if os.path.isfile(filename):
    if filename.endswith('.csv'):
        csv_to_xlsx = csv2xlsx.Convert2xlsx(filename, is_csv=1)
        print 'csv file converted to xlsx extension'
        # convert_into_importable.Convert2importable(csv2xlsx)

    elif filename.endswith('.xls'):
        conv_2_xlsx = csv2xlsx.Convert2xlsx(filename, is_xls=1)
        print 'xls file converted to xlsx extension'
        # convert_into_importable.Convert2importable(conv_2_xlsx)

    elif filename.endswith('.xlsx'):
        con_2_importable = convert_into_importable.Convert2importable(filename)
else:
    print 'Choose a file'
