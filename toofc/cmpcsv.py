from itertools import izip
import os,xlsxwriter
#os.chdir("c://users//nikil//desktop//ntest")

def comp_csv(f1,f2):
    work_dir = f2[:f2.rfind('/')].replace('/','//')
    os.chdir(work_dir)
    w1n = f1[f1.rfind('/')+1:f1.rfind('.')]
    w2n = f2[f2.rfind('/')+1:f2.rfind('.')]
    workbook = xlsxwriter.Workbook(w1n+"vs"+w2n+str('.xlsx'), {'constant_memory': True})
    worksheet = workbook.add_worksheet()
    formt = workbook.add_format()
    formt1 = workbook.add_format()
    formt.set_bg_color('orange')
    formt1.set_bg_color('magenta')
    
    fa,fb = open(f1),open(f2)
    xlrow = 1
    crow=1
    for x,y in izip(fa,fb):
        if x != y:
            worksheet.write(xlrow-1,0,"At line: "+str(crow),formt1)
            #worksheet.set_column(xlrow-1,0,30)
            a,b = x.split(','),y.split(',')
            #print len(a),len(b)
            for j in xrange(len(a)):
                if len(a) <= 1:
                    for j in xrange(len(b)):
                        worksheet.write(xlrow,j,'-')
                else:
                    worksheet.write(xlrow,j,a[j])
                
            for k in xrange(len(b)):
                if len(b) <= 1:
                    for k in xrange(len(a)):
                        worksheet.write(xlrow+1,k,'-')
                else:
                    worksheet.write(xlrow+1,k,b[k])
                
            #print "\nat line "+ str(crow) +" in csv file"
            #print "\r"+ x +"\r" + y
            d = map(None,a,b)
            dl = [i for i in enumerate(d)]
            for e in dl:
                if e[1][0] != e[1][1]:
                    col =  e[0]
                    #print str(xlrow) +"  "+ str(col)+ "  " + str(e)
                    worksheet.write(xlrow+1,col,e[1][1],formt)
            xlrow += 4
        crow += 1
    workbook.close()
    return "created diff workbok successfully"
		
