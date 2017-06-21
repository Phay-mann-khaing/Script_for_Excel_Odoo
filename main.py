import convert2csv
import os
from openpyxl import Workbook

filename = raw_input("File Name: ")
if os.path.isfile(filename):
    if filename.endswith('.csv'):
        test = convert2csv.Convert2csv()
        print 'csv file converted to xlsx extension'
    elif filename.endswith('.xlsx'):
        wb = Workbook()
        ws = wb.active
        ws.append(
            ['line_ids/product_qty', 'line_ids/location_id/id', 'line_ids/product_id/id', 'line_ids/product_uom/id'])
        wb.save("TestCreate.xlsx")
        print "Flie created as 'TestCreate.xlsx'"
    else:
        print 'File type invalid'
else:
    print 'Choose a file'
