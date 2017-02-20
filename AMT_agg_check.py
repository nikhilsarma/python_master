"""
Extract aggregate sum of all the tables given in a list without Multiprocessing
"""

import pyodbc
import xlsxwriter
import os
import time,datetime


os.chdir("target path to store the excel files")
#workbook = xlsxwriter.Workbook('nik_scheck.xlsx', {'constant_memory': True})

workbook = xlsxwriter.Workbook('hsp_mly.xlsx', {'constant_memory': True})
worksheet = workbook.add_worksheet()
bold_italic_f = workbook.add_format({'bold': True, 'italic':True})
bold_f = workbook.add_format({'bold': True})
number_f = workbook.add_format({'num_format': '###,##0.00;[Red](###,##0.00);"-"'})

db1= 'db_uat5'
cnxn_1 = pyodbc.connect('DSN='+db1)
cur1 = cnxn_1.cursor()
print "connection established"

old_q = """
Select NZ_Script from (
Select distinct 'Select Count(1),' as NZ_Script,name from _v_relation_column 
where NAME = 'Table'
union all
Select distinct 'SUM('||ATTNAME||') AS ' ||ATTNAME||',' as NZ_Script, name from _v_relation_column 
where FORMAT_TYPE like 'NUMERIC%'
AND NAME = 'Table'
Union
Select distinct 'SUM('||ATTNAME||') AS ' ||ATTNAME||',' as NZ_Script, name from _v_relation_column 
where FORMAT_TYPE like 'DOUBLE PRECISION'
AND NAME = 'Table'
Union all
Select distinct 'From '||NAME||';' as NZ_Script, name from _v_relation_column 
where NAME = 'Table'
) AS A order by A.NZ_Script DESC;
"""


tb_list = ['list of tables']

tosql_qry = []
def final(query):
    x = cur1.execute(query)
    a = x.fetchall()
    f_qry = ''
    for e in range(len(a)):
        if e == len(a)-2:
            f_qry = f_qry + a[e][0].replace(',',' ')
        else:
            f_qry = f_qry + a[e][0]
    tosql_qry.append(f_qry)
    return f_qry

def todb(q):
        try:
            xyz = cur1.execute(q)
            columns = [t[0] for t in xyz.description]
            result = xyz.fetchall()
            #print "hey"
            #print [columns,result]
        except Exception as e:
            #print "nikhil printing errors"
            print e,e.message
            print q
            columns,result = ['Error'], [(e.message, )]
        #finally:
            #return ['error',0]
        return [columns,result]
    

def new_q(s):
    return old_q.replace("NAME = 'Table'", "NAME = '" + str(s)+"'")

row = 0
for tname in tb_list:
    qry1= new_q(tname)
    qry2 = final(qry1)
    res = todb(qry2)
    #print res[0]
    print tname
    worksheet.write(row,0, tname, bold_italic_f)
    for e in enumerate(res[0]):
        worksheet.write(row+1,e[0],e[1],bold_f)
        #print "i am here nikhil"
    for e in enumerate(res[1][0]):
        #print "i am now here fellow"
        worksheet.write(row+2,e[0],e[1])
    row += 4
   
workbook.close()
