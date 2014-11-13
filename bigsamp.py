from openpyxl import Workbook
wb = Workbook(write_only=True)
ws = wb.create_sheet()

# now we'll fill it with 100 rows x 200 columns

for irow in xrange(10000):
    ws.append(['nikhil' for i in range(10)])
 # save the file
wb.save('c:/users/nikil/desktop/new_big_file.xlsx') 
