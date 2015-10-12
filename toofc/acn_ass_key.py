# longest substring

curString = s[0]
longest = s[0]
for i in range(1, len(s)):
    if s[i] >= curString[-1]:
        curString += s[i]
        if len(curString) > len(longest):
            longest = curString
    else:
        curString = s[i]
print 'Longest substring in alphabetical order is:', longest



#make trans_list rotator

import string

text = "any random text to deciphered"
one = string.ascii_letters
two = one[2:] + one[:2]
transtab = string.maketrans(one,two)
print text.translate(transtab)

#list rotators
"""
Ever wondered why Python list indexing starts at 0, and why slices like
l[x:y] include the start point but not the end point? It's because
operations like this work out so neatly :-)
"""
def rotate(x, y):
    if len(x) == 0:
        return x
    y = y % len(x) # Normalize y, using modulo - even works for negative y
   
    return x[y:] + x[:y]


a = [0,1,2,3,4,5,6]

def rot3(x,y):
    for i in range(y):
        temp = a[0]
        a = a[1:]
        a.append(temp)



#infinite fibo series with generator
def inf_fib():
    while True:
        yield a[0]
        a[1],a[0] = a[0]+a[1],a[1]
        
a = [0,1]    
aa = foo()
for i in range(6):
    print aa.next(),


#recursive multiplication

def recurMul(a,b):
    if b == 1:
        return a
    else:
        return a + recurMul(a,b-1)

#bubble sort
def bbs(lst):
    swp = True
    while swp:
        swp = False
        for i in range(len(lst)-1):
            if lst[i] > lst[i+1]:
                lst[i],lst[i+1] = lst[i+1],lst[i]
                swp = True
    return lst

#sqrt finding
ep = 0.01
y = 4
g = 0.5
while abs(g*g-y) >= ep:
	g = (guess+(y/guess))/2.0
	print g
