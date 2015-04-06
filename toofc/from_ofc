import pyodbc
import os,xlsxwriter,time

os.chdir("C:\\Users\\n.kadayinti\\Desktop\\nikreg")
cnxn = pyodbc.connect('DSN=db_uat')
cursor = cnxn.cursor()

qry1 = """
SELECT COB_DATE,COPER_ID,PARTY_LEGAL_NAME,MAX(FX_DELTA) AS FX_DELTA ,
MAX(MTM)AS MTM,MAX(COMMODITY_DELTA)AS COMMODITY_DELTA,MAX(CS_DELTA)AS CREDITSPREAD_DELTA,
MAX(EQUITY_DELTA) AS EQUITY_DELTA,
SUM(IR_DELTA-IR_DELTA_PAR_POINT) AS INTEREST_RATE_DELTA,
MAX(ISSUER_EXPOSURE) AS ISSUE_EXPOSURE,MAX(JUMP_TO_DEFAULT) AS JUMP_TO_DEFAULT ,MAX(NOTIONAL) AS NOTIONAL
FROM (
SELECT COB_DATE,COPER_ID,PARTY_LEGAL_NAME,
SUM( CASE WHEN MEASURE_NAME='F/X delta' THEN MEASURE  ELSE 0 END) AS FX_DELTA,
SUM( CASE WHEN MEASURE_NAME='MTM' THEN MEASURE  ELSE 0 END) AS MTM,
SUM( CASE WHEN MEASURE_NAME='commodity delta' THEN MEASURE  ELSE 0 END) AS COMMODITY_DELTA,
SUM( CASE WHEN MEASURE_NAME='credit spread delta' THEN MEASURE  ELSE 0 END) AS CS_DELTA,
SUM( CASE WHEN MEASURE_NAME='equity delta' THEN MEASURE  ELSE 0 END) AS EQUITY_DELTA,
SUM( CASE WHEN MEASURE_NAME='interest rate delta' THEN -1* MEASURE  ELSE 0 END) AS IR_DELTA,
SUM( CASE WHEN MEASURE_NAME='ir delta par point' THEN -1*MEASURE  ELSE 0 END) AS IR_DELTA_PAR_POINT,
SUM( CASE WHEN MEASURE_NAME='issuer exposure' THEN MEASURE  ELSE 0 END) AS ISSUER_EXPOSURE,
SUM( CASE WHEN MEASURE_NAME='jump-to-default' THEN MEASURE  ELSE 0 END) AS JUMP_TO_DEFAULT,
SUM( CASE WHEN MEASURE_NAME='notional' THEN MEASURE  ELSE 0 END) AS NOTIONAL
FROM EVRST_HIST..ZINC_SENSITIVITIES_COPER_LEVEL 
WHERE COB_DATE='31-OCT-2014' 
GROUP BY 1,2,3) A
GROUP BY 1,2,3
ORDER BY 2 limit 50
"""
qry4 = """
SELECT COB_DATE,COPER_ID,PARTY_LEGAL_NAME,MAX(FX_DELTA) AS FX_DELTA ,
MAX(MTM)AS MTM,MAX(COMMODITY_DELTA)AS COMMODITY_DELTA,MAX(CS_DELTA)AS CREDITSPREAD_DELTA,
MAX(EQUITY_DELTA) AS EQUITY_DELTA,
SUM(IR_DELTA-IR_DELTA_PAR_POINT) AS INTEREST_RATE_DELTA,
MAX(ISSUER_EXPOSURE) AS ISSUE_EXPOSURE,MAX(JUMP_TO_DEFAULT) AS JUMP_TO_DEFAULT ,MAX(NOTIONAL) AS NOTIONAL
FROM (
SELECT COB_DATE,COPER_ID,PARTY_LEGAL_NAME,
SUM( CASE WHEN MEASURE_NAME='F/X delta' THEN MEASURE  ELSE 0 END) AS FX_DELTA,
SUM( CASE WHEN MEASURE_NAME='MTM' THEN MEASURE  ELSE 0 END) AS MTM,
SUM( CASE WHEN MEASURE_NAME='commodity delta' THEN MEASURE  ELSE 0 END) AS COMMODITY_DELTA,
SUM( CASE WHEN MEASURE_NAME='credit spread delta' THEN MEASURE  ELSE 0 END) AS CS_DELTA,
SUM( CASE WHEN MEASURE_NAME='equity delta' THEN MEASURE  ELSE 0 END) AS EQUITY_DELTA,
SUM( CASE WHEN MEASURE_NAME='interest rate delta' THEN -1* MEASURE  ELSE 0 END) AS IR_DELTA,
SUM( CASE WHEN MEASURE_NAME='ir delta par point' THEN -1*MEASURE  ELSE 0 END) AS IR_DELTA_PAR_POINT,
SUM( CASE WHEN MEASURE_NAME='issuer exposure' THEN MEASURE  ELSE 0 END) AS ISSUER_EXPOSURE,
SUM( CASE WHEN MEASURE_NAME='jump-to-default' THEN MEASURE  ELSE 0 END) AS JUMP_TO_DEFAULT,
SUM( CASE WHEN MEASURE_NAME='notional' THEN MEASURE  ELSE 0 END) AS NOTIONAL
FROM EVRST_HIST..ZINC_SENSITIVITIES_COPER_LEVEL 
WHERE COB_DATE='31-OCT-2014' 
GROUP BY 1,2,3) A
GROUP BY 1,2,3
ORDER BY 2
"""
qry2 = """
select CUST_ID, CUSTOMER_GCI_NBR, CUSTOMER_NM, UPGCI_FLAG, CED_ENTITY_ID, PRIMARY_NAICS_NOTIONAL_LIMIT_AMT, 
MOODY_EXTERNAL_RATING_DT, CLL_AMT, PRDS_INTCO_COST_CENTER, PERIOD_DT from EVRST_UI_RPT..CUSTOMER_AW_SVT limit 5
"""
qry3 = """ select CUST_ID, CUSTOMER_GCI_NBR, CUSTOMER_NM, UPGCI_FLAG, CED_ENTITY_ID, PRIMARY_NAICS_NOTIONAL_LIMIT_AMT, 
MOODY_EXTERNAL_RATING_DT, CLL_AMT, PRDS_INTCO_COST_CENTER, PERIOD_DT from EVRST_UI_RPT..CUSTOMER_AW_SVT
where CLL_AMT is not NULL and CLL_AMT != 0.00"""

ab = cursor.execute(qry4)
workbook = xlsxwriter.Workbook('dbtest_new_new.xlsx', {'constant_memory': True})
worksheet = workbook.add_worksheet("nikrest")
bold_f = workbook.add_format({'bold': True})
italic_f = workbook.add_format({'italic': True})
number_f = workbook.add_format({'num_format': '###,##0.00;[Red](###,##0.00);"-"'})
string_f = workbook.add_format({'num_format': 'General'})
date_f = workbook.add_format({'num_format': 'd-mmm-yy'})


num_stack = ['int','long','decimal']
date_stack = ['date']
columns = [t[0] for t in ab.description]
for e in enumerate(ab.description):
        #print e[0],e[1][0],e[1][1]
        #worksheet.write(0,e[0],e[1][0],bold_f)
        #worksheet.set_row(0, None, bold_f)
        #if e[1][1].__name__ in num_stack:
                #worksheet.set_column(e[0],e[0], None, number_f)
                #worksheet.write(0,e[0],e[1][0],bold_f)
                #print "hai"
        if e[1][1].__name__ in date_stack:
                worksheet.set_column(e[0],e[0], None, date_f)
                worksheet.write(0,e[0],e[1][0],bold_f)
                #print "hello"
        else:
                worksheet.set_column(e[0],e[0], None, string_f)
                worksheet.write(0,e[0],e[1][0],bold_f)
                #print "pollo"
        #print e,
        #worksheet.write(0,e[0],e[1][0],bold_f)
        
r,c = 1,0
flag = 1
t1 = time.time()
#print t1
while flag:
        #print "nikhil"
        try:
                for yy in enumerate(ab.fetchone()):
                        #print yy
                        #data = unicode(yy[1])
                        #print data
                        worksheet.write(r,yy[0],yy[1])
                        #print r, yy[0],yy[1]
                r+= 1
        except:
                flag = 0
workbook.close()
t2 = time.time()
print round((t2-t1),2)
