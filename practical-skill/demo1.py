"""
@author:maohui
@time:2022/3/2 16:17
"""


# 1.重复元素判断:检查给定的列表是否存在重复的元素
def all_unique(list):
    return len(list) == len(set(list))


# 2.字符元素组成判定：检查两个字符串的组成元素是不是一样的
from collections import Counter


def anagram(first, second):
    return Counter(first) == Counter(second)


# 3.内存占用：检查变量variable所占用的内存(字节)
import sys


def memory(variable):
    return sys.getsizeof(variable)


# 4.字节占用：检查字符串占用的字节数
def byte_size(string):
    return (len(string.encode('utf-8')))


# 5.打印n次方字符串：不使用循环
def ptloop(str, n):
    print(str * n)


# 6.大写第一个字母:大写字符串中的每一个单词的首字母
def title_cap(s):
    print(s.title())


# 7.分块：按照具体的大小切割列表
from math import ceil


def chunk(listin, size):
    return list(map(lambda x: listin[x * size:x * size + size], list(range(0, ceil(len(listin) / size)))))


# 8.压缩：使用filter（）函数去掉布尔型的值（例：False，None，0，“”）,不包括True
def compact(listin):
    return list(filter(bool, listin))


# 9.解包：将打包好的成对列表按找index解开成两组不同的元组
def unpack(arrayin):
    transposed = zip(*arrayin)
    list1 = []
    for i in transposed:
        list1.append(i)
    print(list1)


# 10.链式对比:在一行代码使用不同的运算符对比多个不同的元素
def compare():
    a = 3
    print(1 < a < 8)
    print(1 == a < 4)


# 11.逗号连接
def comma():
    hobbies = ['basketball', 'football', 'swimming']
    print('my hobbies are:' + ', '.join(hobbies))


# 12.元音统计：通过正则表达式统计（a,e,i,o,u）的个数
import re


def count_vowels(str):
    return len(re.findall(r'[aeiou]', str, re.IGNORECASE))

#13.首字母小写：字符串的第一个字符小写
def decapitalize(str):
    return str[:1].lower()+str[1:]
if __name__ == "__main__":
    list1 = [1, 2, 3, 1]
    list2 = [1, 2, 3]
    list3 = [2, 1, 3]
    print('---------1.------------')
    print(all_unique(list1))
    print(all_unique(list2))
    print('---------2.-------------')
    print(anagram(list1, list2))
    print(anagram(list3, list2))
    print(anagram('12', '21'))
    print('---------3.-------------')
    print(memory(30))
    print('---------4.-------------')
    print(byte_size('hello world'))
    print(byte_size('Hello World'))
    print('---------5.-------------')
    ptloop('programming', 3)
    print('---------6.-------------')
    title_cap('programming is awesome')
    print('---------7.-------------')
    print(chunk([1, 2, 3, 4, 5], 2))
    print('---------8.-------------')
    print(compact([0, 1, False, True, 2, "", "a", 3, None]))
    print('---------9.-------------')
    array = (['a', 'b'], ['c', 'd'], ['e', 'f'])
    unpack(array)
    print('---------10.-------------')
    compare()
    print('---------11.-------------')
    comma()
    print('---------12.-------------')
    print(count_vowels('foobar'))
    print('---------13.-------------')
    print(decapitalize('Foobal'))