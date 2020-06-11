from django.test import TestCase

# Create your tests here.

# s = '黑龙江'
# n = ['北京1','北京2','黑龙江1','黑龙江2']
# n = ['北京3','北京4','黑龙江5','黑龙江6']
#
# b = ['北京','北京','黑龙江','黑龙江']
#
# T = sorted(set(b), key=b.index)
# print(T)
# c = [i[:2] + str(j + 1) for i in T for j in range(b.count(i))]
# print(c)
#
# nton = dict(zip(c,n))
# print(nton)
#
# s = '北京-上海'
# print(s.split('-'))
a = {
     'x': 1,
     'y': 2,
     'z': 3
}

b = {
     'w': 10,
     'x': 11,
     'y': 2
}

c = {
     'w': 11,
     'x': 10,
     'y': 2
}

print('Common keys:', a.keys() & b.keys())
print('Keys in a not in b:', a.keys() - b.keys())
print('Keys in a or in b', a.keys() | b.keys())

print('(key,value) pairs in common:', a.items() & b.items())
print('(key,value) pairs in a not in b:', a.items() - b.items())
print('(key,value) pairs in a or in b:', a.items() | b.items())