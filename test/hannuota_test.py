"""
@author:maohui
@time:2022/3/2 10:24
"""

def move(n,a,b,c):
    if n==1:
        print(a, '>', c)
    else:
        move(n-1,a,c,b)
        move(1,a,b,c)
        move(n-1,b,a,c)
move(3,'a','b','c')