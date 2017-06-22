import convert2csv
import os
import openpyxl

filename = raw_input("File Name: ")
if os.path.isfile(filename):
    if filename.endswith('.csv'):
        test = convert2csv.Convert2csv(filename)
        print 'csv file converted to xlsx extension'
    elif filename.endswith('.xlsx'):
        wb = openpyxl.load_workbook(filename)
        wb1 = openpyxl.Workbook()
        ws = wb.active
        ws1 = wb1.active
        column_count = ws.max_column - 3
        row_count = ws.max_row

        a = list(ws.iter_rows())
        for index1, line in enumerate(a):
            for index2, piece in enumerate(line):
                print piece.value

        col_a = ws['A']  # 0-indexing
        for idx, cell in enumerate(col_a, 1):
            ws1.cell(row=idx, column=1).value = cell.value

        # for cell in col_a:
        # ws.cell(row=idx, column=4).value = cell.value # 1-indexing


        # ws1.append(
        #     ['line_ids/product_qty', 'line_ids/location_id/id', 'line_ids/product_id/id', 'line_ids/product_uom/id'])

        ws1['A1'] = 'line_ids/product_id/id'

        wb1.save("/home/phay/PycharmProjects/TestCreate.xlsx")
        print "Flie created as 'TestCreate.xlsx'"
    else:
        print 'File type invalid'
else:
    print 'Choose a file'
