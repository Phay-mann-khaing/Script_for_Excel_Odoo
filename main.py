import convert2csv
import os
import openpyxl

filename = raw_input("File Name: ")
if os.path.isfile(filename):
    if filename.endswith('.csv'):
        test = convert2csv.Convert2csv()
    elif filename.endswith('.xlsx'):
        wb = workbook()
        pass
    else:
        print 'File type invalid'
else:
    print 'Choose a file'
