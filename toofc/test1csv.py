from itertools import izip_longest
import os,xlsxwriter,time
#os.chdir("c://users//nikil//desktop//ntest")

coldic = {'Cocoa':'D2691E','Aqua':'7EC0EE','Olive':'89892B','Yellow':'CCA300','Orange':'FF8000'} 

def the_mess(l):
    #col,eps = l[3],l[4]
    print l
    print "Comparision started...!"
    
    report = []
    if l[1] != '' and l[2] != '':
        f1,f2 = l[1],l[2]
        work_dir = l[2][:l[2].rfind('/')].replace('/','//')
        #print work_dir
        os.chdir(work_dir)
        rprt = comp_csv(f1,f2)
        report.append(rprt)
    elif l[0] != '':
        work_dir = l[0].replace("/","//")
        os.chdir(work_dir)
        files_list = [e.lower() for e in filter(lambda x: x.endswith('.csv'), os.listdir(work_dir))]
        print files_list.sort()
        lfile = len(files_list)
        cnt = 0
        while cnt < lfile:
            f1 = files_list[cnt]
            f2 = files_list[cnt+1]
            cnt += 2
            rprt = comp_csv(f1,f2)
            report.append(rprt)
    return report

def itercount(filename):
    return sum(1 for _ in open(filename, 'rbU'))

def comp_csv(f1,f2):

    f1c,f2c = itercount(f1),itercount(f2)
    #print "hi"
    print "Comparing... " +f1 +" : "+ str(f1c) + " vs " + f2 +" : "+ str(f2c) 
    t1 = time.time()
    w1n = f1[f1.rfind('/')+1:f1.rfind('.')]
    w2n = f2[f2.rfind('/')+1:f2.rfind('.')]
    workbook = xlsxwriter.Workbook(w1n+"_vs_"+w2n+str('.xlsx'), {'constant_memory': True})
    worksheet = workbook.add_worksheet()
    formt = workbook.add_format()
    formt1 = workbook.add_format()
    formt.set_bg_color('orange')
    formt1.set_bg_color('magenta')
    sucmsg = "Pass"
    fa,fb = open(f1),open(f2)
    xlrow = 1
    crow=1
    for x,y in izip_longest(fa,fb):
        
        if x != y:
            if x == None:
                x = ''
            elif y == None:
                y = ''
            sucmsg = "Fail"
            worksheet.write(xlrow-1,0,"At line: "+str(crow),formt1)
            #worksheet.set_column(xlrow-1,0,30)
            a,b = x.strip().split(','),y.strip().split(',')
            alen,blen = len(a),len(b)
            #print len(a),len(b)
            for j in xrange(alen):
                if alen <= 1:
                    for j in xrange(blen):
                        worksheet.write(xlrow,j,'-')
                else:
                    worksheet.write(xlrow,j,a[j])
                
            for k in xrange(blen):
                if blen <= 1:
                    for k in xrange(alen):
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
                    print str(xlrow) +"  "+ str(col)+ "  " + str(e)
                    worksheet.write(xlrow+1,col,e[1][1],formt)
            xlrow += 4
        crow += 1
    workbook.close()
    t2 = time.time()
    td = round((t2-t1),2)
    return [td,sucmsg,w1n+" vs "+w2n,w1n+"_vs_"+w2n+str('.xlsx')]
    #return "created diff workbok '" + w1n+"_vs_"+w2n+str('.xlsx')+"' successfully in :" + str(td) + "sec..."
		



