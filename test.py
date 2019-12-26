# field_type={
#     "n":'1',
#     "s":1,
#     "t":1,
#     "an":1,
#     "as":1,
#     "at":1,
#     "d":1
# }
# key_type='all_mainkey'
# import data_checker
# name='n_ddd'
# data_checker.field_name([1,2],field_type,name)
# data_checker.field_range([1,2],key_type)
# print('hhah')
# import os
# a=os.walk('D:\\7zcracker')
# for file in a:
#     print(file)
# for i in range(1,5):
#     print(i)
# a=(5,10)
# for i in range(a[0],a[1]):
#     print(i)
import xlrd
a=xlrd.open_workbook('.\\test.xlsx').sheet_by_name("test")
b=a.cell(5,0).value
print(a.ncols)