import datetime, time
from openpyxl import load_workbook
f1 = "c:/users/nikil/desktop/bf1.xlsx"
f2 = "c:/users/nikil/desktop/bf2.xlsx"
t1 = time.time()
wb1 = load_workbook(f1,use_iterators = True)
wb2 = load_workbook(f2,use_iterators = True)
ws1 = wb1.worksheets[0]
ws2 = wb2.worksheets[0]
r1 = ws1.iter_rows()
r2 = ws2.iter_rows()
cnt = 1
while cnt <= ws1.get_highest_row():
    rone = r1.next()
    rtwo = r2.next()
    for i in xrange(ws1.get_highest_column()):
        if rone[i].value != rtwo[i].value:
            print "mismatch found at row: " +str(cnt)+" in cell: " +str(i+1)+"."
            print str(rone[i].value)+ " ---> " + str(rtwo[i].value)
    cnt += 1


t2 = time.time()
ttime = round((t2-t1),2)
print "total time of execution is: " +str(ttime) + "sec."

"""		
	except StopIteration:
		print "end of comp"
		t2 = time.time()
		ttime = round((t2-t1),2)
		print "total time of execution is: " +str(ttime) + "sec."
		break
"""
