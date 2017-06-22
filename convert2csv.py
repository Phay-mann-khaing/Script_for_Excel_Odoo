import csv
import openpyxl

class Convert2csv:
    def __init__(self, filename=None):
        self.filename = filename
        self.convert2csv()

    def convert2csv(self):
        try:
            wb = openpyxl.Workbook()
            ws = wb.active

            f = open(self.filename)
            reader = csv.reader(f, delimiter=',')
            for row in reader:
                # for list_index in row:
                #     words = list_index.split(",")
                ws.append(row)

            f.close()

            wb.save('/home/phay/PycharmProjects/Convertedtoxlsx.xlsx')

        except ValueError:
            print 'File invalid'
            # except CustomException, (instance):
            #     self.ok = False

# class CustomException(Exception):
#     def __init__(self, value):
#         self.parameter = value
#
#     def __str__(self):
#         return repr(self.parameter)

# test_convert = Convert2csv()
