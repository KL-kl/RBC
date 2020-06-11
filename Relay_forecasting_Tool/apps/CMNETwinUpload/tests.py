# import xlrd
# import os
# import time
# import random
# import json
# import inspect
# from functools import reduce
#
# print(time.time())
# # 格式化时间戳为标准格式
# print(time.strftime('%Y%m%d%H%M%S', time.localtime(time.time())))
# print(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
# #
# List = [1, 2, 2, 2, 2, 3, 3, 3, 4, 4, 4, 4]
#
# b = []
# for i in List:
#     a = {}
#
#     if List.count(i) > 1:
#         a[i] = List.count(i)
#
# # print(a,len(a))
# # for i in range(len(a)):
# #     print(i)
#
# # j = json.dumps(a)
# # print(list(j))
# data = ({'province': '北京'}, {'province': '北京'}, {'province': '上海'},{'province': '上海'},{'province': '上海'}, {'province': '上海'}, {'province': '广东'}, {'province': '广东'})
#
# prolist = [i['province'] for i in data]
#
# new = reduce(lambda x, y: x if y in x else x + [y], [[], ] + list(data))
#
# for i in new:
#
#     i['number'] = prolist.count(i['province'])
#
# print(new)
# a = list(filter(lambda x: x['number'] > 2, new))
# print(a)
#
# str = ""
# for i in range(4):
#     ch = chr(random.randrange(ord('0'), ord('9') + 1)) # ord(char)函数将char类型的单字符转换成ASCII码值
#
#     str += ch
#
# str = time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))+str
import re


# n = s.split('~')[:2]
# m = s[:2]
print(res)