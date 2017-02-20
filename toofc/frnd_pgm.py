import os
from datetime import datetime
from xlwt import Workbook

os.chdir("C:/Users/nikil/Desktop/frnd_pgms")
fl = [e.lower() for e in filter(lambda x: x.endswith('.log') ,os.listdir(os.getcwd()))]

def foo(d_string):
	q = []
	for i in d_string.split(','):q.append(i.strip(" '"))
	q.sort(key=lambda date: datetime.strptime(date, "%d %B %Y").strftime("%d-%b-%y"), reverse = True)
	return q

def bar(f):
	fa = open(f)
	res = {}
	res['name'] = f[:f.rfind('.')]
	txt = fa.read()
	df = txt.find("Dates found")
	dm = txt.find("Dates min")
	if df > 1:
		st = txt.find('[',df)
		ed = txt.find(']',st)
		res['Dates found'] = foo(txt[st+1:ed])
	else:
		res['Dates found'] = [None]
	if dm > 1:
		st = txt.find('[',dm)
		ed = txt.find(']',st)
		res['Dates min'] = foo(txt[st+1:ed])
	else:
		res['Dates min'] = [None]
	#print res
	return res

book = Workbook(encoding='utf-8')
s1 = book.add_sheet("Dates Range")
s2 = book.add_sheet("Dates count")
header = ['Name','Dates found','Dates min']
for i in enumerate(header):
    s1.write(0,i[0],i[1])
    s2.write(0,i[0],i[1])
ofset_s1 = 1
ofset_s2 = 1
for files in fl:
    oput = bar(files)
    print oput
    result = map(None,oput['Dates found'],oput['Dates min'])
    df = None if oput['Dates found'] == [None] else len(oput['Dates found'])
    dm = None if oput['Dates min'] == [None] else len(oput['Dates min'])
    counts = [oput['name'],df,dm]
    s1.write(ofset_s1,0,files)
    #s1.merge(ofset_s1,counts[1],0,0)
    for e in enumerate(counts):
        s2.write(ofset_s2,e[0],e[1])
    ofset_s2 += 1
    for e in enumerate(result,ofset_s1):
        s1.write(e[0],1,e[1][0])
        s1.write(e[0],2,e[1][1])
        ofset_s1 = e[0]+2
        
book.save("final.xls")
