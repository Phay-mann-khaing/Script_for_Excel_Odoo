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
        column_count = ws.max_column
        row_count = 1
        max_current_row = 2
        a = list(ws.iter_rows())
        for index1, line in enumerate(a, 1):
            for index2, piece in enumerate(line, 1):
                if index1 == 1:
                    ws1.cell(row=index1, column=index2).value = piece.value
                else:
                    count = column_count - 3
                    row = 0
                    row = row_count

                    while count > 0:
                        ws1.cell(row=row, column=index2).value = piece.value
                        count -= 1
                        row += 1
                    if max_current_row < row:
                        max_current_row = row

            row_count = max_current_row

        # col_a = ws['A']  # 0-indexing
        # for idx, cell in enumerate(col_a, 1):
        #     ws1.cell(row=idx, column=1).value = cell.value

        # for cell in col_a:
        # ws.cell(row=idx, column=4).value = cell.value # 1-indexing


        # ws1.append(
        #     ['line_ids/product_qty', 'line_ids/location_id/id', 'line_ids/product_id/id', 'line_ids/product_uom/id'])

        # ws1['A1'] = 'line_ids/product_id/id'

        wb1.save("/home/phay/PycharmProjects/TestCreate.xlsx")
        print "Flie created as 'TestCreate.xlsx'"
    else:
        print 'File type invalid'
else:
    print 'Choose a file'
