#import pyodbc
import os,xlsxwriter,time,datetime
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Style, fills, PatternFill, Color

#os.chdir("C:\\Users\\n.kadayinti\\Desktop\\nikreg")

#db1,db2 = "db_uat","db_prod"
#cnxn_1,cnxn_2 = pyodbc.connect('DSN='+db1), pyodbc.connect('DSN='+db2)
#cur1,cur2= cnxn_1.cursor(),cnxn_2.cursor()

#cnxn_uat,cnxn_prod = pyodbc.connect('DSN=db_uat'), pyodbc.connect('DSN=db_prod')
#c_uat,c_prod = cnxn_uat.cursor(),cnxn_prod.cursor()


#l = ['TC004', 'F_COLLATERAL_CountCheck', "select count(*) from EVRST_HIST_NK..CLC_F_COLLATERAL where period_dt='2014-12-11'"]

coldic = {'Cocoa':'D2691E','Blue':'7EC0EE','Olive':'89892B','Yellow':'CCA300','Orange':'FF8000','Rose':'FF7C80','Green':'99CC00'}



def compare(s1, s2, w1n, w2n):

    rowrange = max(s1.max_row, s2.max_row)
    colrange = max(s1.max_column, s2.max_column)
    cfill1 = Style(fill=PatternFill(patternType='solid', fgColor=Color(coldic["Yellow"])))
    cfill2 = Style(fill=PatternFill(patternType='solid', fgColor=Color(coldic["Green"])))
    cfill3 = Style(fill=PatternFill(patternType='solid', fgColor=Color(coldic["Rose"])))
    succ = "Pass"
    e = 'Data Matches!'
    try:
        for i in xrange(1,rowrange+1):
            for j in xrange(1,colrange+1):
                #time.sleep(0.5)
                x = s1.cell(row=i,column=j)
                y = s2.cell(row=i,column=j)
                xfor_code,yfor_code = x.number_format,y.number_format
                
                a = x.value
                b = y.value
                nt = None
                dt = datetime.datetime.today()
                tlist = [type(int()), type(float()),type(long()), type(nt)]
                numlist = [type(int()), type(float()),type(long())]
                str_list = ["d-mmm-yy","@"]
                #xfor_code = s2.cell(row=i,column=j).number_format
                if (type(a) not in numlist or x.number_format in str_list) and (type(b) not in numlist or y.number_format in str_list):
                    a,b = unicode(a).lower(),unicode(b).lower()
                    #print x.number_format+str('->')+y.number_format
                    #print a + str('->')+ b
                    #print str(type(a)) +str('->')+str(type(b))
                if a != b:
                    #print a,b
                    #print type(a),type(b)
                    succ = "Fail"
                    e = "Data Mismatches"
                    comtxt = None
                    
                    if type(a) in numlist and (type(b) == type(nt) or type(b) == type(unicode())):
                        comtxt = str(w1n)+": " + unicode(a)+ str("\ndiff: ")+ unicode(a)
                    
                    elif type(b) in numlist and (type(a) == type(nt) or type(a) == type(unicode())):
                        comtxt = str(w1n)+": " + unicode(a)+ str("\ndiff: ")+ unicode(b)
                    
                    elif type(a) in numlist and type(b) in numlist:
                        diff_val = float(abs(a-b))
                        y.style = cfill1 if diff_val <= float(0.009) else cfill2 if diff_val <= float(100) else cfill3
                        #if diff_val <= float(0.001):
                            #print "ignored" + str(a) + str(" and ") +str(b)
                            #continue
                        if a < b:
                            comtxt = str(w1n)+": " + unicode(a)+ str("\ndiff: ")+ unicode(b-a)+ str("\n")+unicode(round(diff_val,4))+ unicode(" % inc. " )
                        elif a > b:
                            comtxt = str(w1n)+": " + unicode(a)+ str("\ndiff: ")+ unicode(b-a)+str("\n")+unicode(round(diff_val,4))+ unicode(" % dec. " )
                        #comtxt = str(w1n)+": " + unicode(a)+ str("\ndiff: ")+ unicode(b-a)+ str("\ndiff %: ")+ unicode(diff_per)

                    elif type(a) == type(unicode()) and type(b) == type(unicode()):
                        comtxt = str(w1n)+": " + unicode(a)
                
                    comment = Comment(comtxt, w2n)
                    y.style = cfill1
                    y.comment = comment
                    y.number_format = yfor_code
                else:
                    comtxt = None
    except KeyboardInterrupt:
        succ = "Abort"
        e = "KeyboardInterrupt"
    except Exception as e:
        succ = "Abort"
            
    #print [succ,e]
    return [succ,e]


def toexport(i,cur,qry,tc_no,tc,db1,db2):
    #print i
    cur_dir = os.getcwd()
    if not os.path.exists(tc_no):
        os.makedirs(tc_no)
    #os.makedirs(tc_no)
    folname = os.path.join(os.getcwd(),tc_no)
    filname = tc_no + "_" +  (db1 if i == 0 else db2)
    #print folname
    workbook = xlsxwriter.Workbook(folname+"\\" + filname +".xlsx", {'constant_memory': True})
    worksheet = workbook.add_worksheet()
    bold_f = workbook.add_format({'bold': True})
    string_f = workbook.add_format({'num_format': 'General'})
    date_f = workbook.add_format({'num_format': 'd-mmm-yy'})
    italic_f = workbook.add_format({'italic': True})
    number_f = workbook.add_format({'num_format': '###,##0.00;[Red](###,##0.00);"-"'})
    num_stack = ['int','long','decimal']
    date_stack = ['date']
    try:
        #print qry
        x = cur.execute(qry)
        columns = [t[0] for t in x.description]
        #msg = "Pass"
	#ex_stat = "Y"
        #print columns
        for e in enumerate(x.description):
            if e[1][1].__name__ in date_stack:
                worksheet.set_column(e[0],e[0], None, date_f)
                worksheet.write(0,e[0],e[1][0],bold_f)
                #print "hello"
            else:
                worksheet.set_column(e[0],e[0], None, string_f)
                worksheet.write(0,e[0],e[1][0],bold_f)
        r,c = 1,0
        flag = 1
        while flag:
            try:
                for yy in enumerate(x.fetchone()):
                    worksheet.write(r,yy[0],yy[1])
                r += 1
            except:
                flag = 0
        row_count = r-1
    except Exception as e:
        msg = "Fail"
	ex_stat = "N"
    workbook.close()
    return row_count
   
  
def the_mayhem(f1,f2):
    print "Comparing... " +f1 +" vs " + f2
    t1 = time.time()
    w1 = load_workbook(f1)
    w2 = load_workbook(f2)
    w1n = f1[f1.rfind('/')+1:f1.rfind('.')]
    w2n = f2[f2.rfind('/')+1:f2.rfind('.')]
    wbsucc = "Pass"
    sdic = {}

    for i in xrange(len(w1.worksheets)):
        for j in xrange(i,i+1):
            smsg = compare(w1.worksheets[j], w2.worksheets[j],w1n,w2n)
            sdic[w1.worksheets[j].title] = smsg[0],smsg[1]
            w2.save(f2)
    for e in sdic.values():
        if "Pass" not in e[0]:
            wbsucc = "Fail"
    
    t2 = time.time()
    ttime = round((t2-t1),2)
    #print sdic
    return [ttime,w1n,w2n,wbsucc,sdic]

def the_mess(cl,comp_dir):
    #print cl
    print "File Comparision started...! at " + str(comp_dir) + "..."
    report = []
    if cl[0] != '' and cl[1] != '':
        f1 = os.path.join(comp_dir,cl[0]).replace("\\","/")
        f2 = os.path.join(comp_dir,cl[1]).replace("\\","/")
        rprt = the_mayhem(f1,f2)
        report.append(rprt)
    return report,f2


def all_check(tc_no,tc,c_qry,cur1,cur2,db1,db2):
    #print tc_no,tc,c_qry,cur1,cur2,db1,db2
    rcount = []
    for e in enumerate([cur1,cur2]):
        #print e[0],e[1]
        r_c = toexport(e[0],e[1],c_qry,tc_no,tc,db1,db2)
        rcount.append(r_c)
    #print rcount
    comp_dir = os.path.join(os.getcwd(),tc_no)
    tocon_list = filter(lambda x: x.endswith('.xlsx'), os.listdir(comp_dir))
    #print tocon_list
    res = the_mess(tocon_list,comp_dir)
    #print res
    return res,rcount



