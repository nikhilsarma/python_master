"""

def initialise(a,b):
    cnxn_uat,cnxn_prod = pyodbc.connect('DSN=db_uat'), pyodbc.connect('DSN=db_prod')

    if 

"""

import pyodbc
import os,xlsxwriter,time
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Style, fills, PatternFill, Color
from exp_final import *

def count_check(tc_no,tc,c_qry,cur1,cur2):
	#print"\n"
	#print c_qry
	import time
	try:
		t1 = time.time()
		c1,c2 = cur1.execute(c_qry),cur2.execute(c_qry)
		a,b = c1.fetchone(),c2.fetchone()
		a1 = 0 if a == None else a[0]
		b1 = 0 if b == None else b[0]
		#print a[0],b[0]
		t2 = time.time()
		tt = round((t2-t1),2)
	except Exception as e:
		msg = "Fail"
		ex_stat = "N"
		
		return [tc_no,tc,ex_stat,msg,(e.__class__.__name__,e.args),str(round((time.time()-t1),2))]
	else:
		msg = "Pass" if a == b else "Fail"			
		ex_stat = "Y"		
		return [tc_no,tc,ex_stat,msg,(a1,b1),str(tt)]


def runto_excel():
    
    db1,db2 = "db_uat","db_prod"
    cnxn_1,cnxn_2 = pyodbc.connect('DSN='+db1), pyodbc.connect('DSN='+db2)
    cur1,cur2= cnxn_1.cursor(),cnxn_2.cursor()
    #cnxn_uat,cnxn_prod = pyodbc.connect('DSN=db_uat'), pyodbc.connect('DSN=db_prod')
    #c_uat,c_prod = cnxn_uat.cursor(),cnxn_prod.cursor()
    os.chdir("C:\\Users\\n.kadayinti\\Desktop\\nikreg")
    f1 = "c:/users/n.kadayinti/desktop/nikreg/qry_test.xlsx"
    wb1 = load_workbook(f1,use_iterators = True)
    workbook = xlsxwriter.Workbook('dbtest_count_check.xlsx', {'constant_memory': True})
    bold_f = workbook.add_format({'bold': True})
    for e in wb1.worksheets:
        ws = e
        row = ws.iter_rows()
        row.next()
        s_name = ws.title
        worksheet = workbook.add_worksheet(s_name)
        xl_hd = ["Test_Case_No","Test_Condition","SQL Query","Executed","Result","O/P: Comments","Time Taken"]
        for e in enumerate(xl_hd):
            worksheet.write(0,e[0],e[1],bold_f)
        for e in enumerate(row):
            if e[1][1].value[-10:] == "CountCheck":
                #res = count_check(e[0].value,e[1].value,e[2].value.replace('\n',''))
                res = count_check(e[1][0].value,e[1][1].value,e[1][2].value.replace('\n',' '),cur1,cur2)
                res.insert(2,e[1][2].value)
                print res
                for col in enumerate(res):
                    #print col
                    if col[0] == 5:
                        if type(col[1][0]) and type(col[1][1]) == type(1l):
                            cmt = db1+":" + str(col[1][0]) + " ; " + db2+":" + str(col[1][1])
                        else:
                            cmt = str(col[1][0]) +" : " + str(col[1][1][1].replace(' ',''))
                    worksheet.write(e[0]+1,col[0],cmt if col[0] == 5 else col[1])
                    #worksheet.write(e[0],col[0],e[1][2].value if col[0] == 3 else col[1])
            else:
                    res = all_check(e[1][0].value,e[1][1].value,e[1][2].value.replace('\n',' '),cur1,cur2,db1,db2)
                    #print "now in else part of runtoexcel"
                    for col in enumerate(res):
                            if col[0] == 5:
                                    if type(col[1][0]) and type(col[1][1]) in [type(1l),type(2)]:
                                            cmt = db1+":" + str(col[1][0]) + " ; " + db2+":" + str(col[1][1])
                                            
                    print res
                    
                        
                    
