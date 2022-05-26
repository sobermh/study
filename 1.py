"""
@author:maohui
@time:2022/5/25 11:30
"""

cout=1
i=int(input(">>>"))
def er(cout,i):
    while i>10:
        cout+=1
        i-=1
        err(cout)
        print(cout)
def err(cout):
    cout-=1
    return cout
er(cout,i)
