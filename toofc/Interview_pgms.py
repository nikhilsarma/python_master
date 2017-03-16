global monthDict
monthDict={'Jan': (1,31),'Feb':(2,28,29),'Mar': (3,31),'Apr': (4,30),'May':(5,31),'Jun': (6,30),'Jul': (7,31),'Aug':(8,31),'Sep': (9,30),'Oct': (10,31),'Nov':(11,30),'Dec':(12,31)}

def time_elapsed(a,b):
	d1,m1,y1 = a.split('-')
	d1,m1,y1 = int(d1),m1,int(y1)
	d2,m2,y2 = b.split('-')
	d2,m2,y2 = int(d2),m2,int(y2)

	if y1 == y2:
		if m1 == m2:
			if d1 == d2:
				return "0 Days!"
			return str(abs(d1-d2)) + " Days!"
		elif monthDict[m1][0] > monthDict[m2][0]:
			return remdays(y1,f = (m2,d2), to = (m1,d1)) 
		elif monthDict[m2][0] > monthDict[m1][0]:
			return remdays(y1,f = (m1,d1), to = (m2,d2)) 
	elif y1 > y2:
		raise Exception ("Choose form same year")
	elif y1 < y2:
		raise Exception ("Choose from same year")


def remdays(y, **kwargs):
	year = y
	fm,fd = kwargs['f']
	tm,td = kwargs['to']
	remdays_curmonth = monthDict[fm][1] - fd
	remmonths,remdays = 0,remdays_curmonth
	for e in monthDict.values():
		if e[0] > monthDict[fm][0] and e[0] <= monthDict[tm][0]:
			remmonths += 1
			if e[0] == monthDict[tm][0]:
				remdays += td
				if fd > td:
					remmonths -= 1
				continue
			remdays += e[1]

	return "{} Months & {} days have Elapsed!!".format(remmonths,abs(fd-td))
	


a = "05-Apr-2017"
b = "01-Mar-2017"

print time_elapsed(a,b)
