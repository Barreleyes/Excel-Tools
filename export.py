import sys
import os
import argparse
# 创建外部参数解析器
parser = argparse.ArgumentParser()

# 添加外部参数参数(positional arguments)
parser.add_argument('-i', '--input', default='.\\',
                    help="excel文件或文件夹路径，默认读取【工具目录】")
parser.add_argument('-oc', '--output-client', default='.\\export\\client\\',
                    help="前端导出文件路径，默认在【工具目录\\export\\client】文件夹")
parser.add_argument('-os', '--output-server',default='.\\export\\server\\',
                    help="后端导出文件路径，默认在【工具目录\\export\\server】文件夹")
parser.add_argument('-tc', '--type-client', default='lua',
                    help="前端导出格式，可根据项目定制，默认为lua")
parser.add_argument('-ts', '--type-server', default='lua',
                    help="后端导出格式，可根据项目定制，默认为lua")
# 外部参数解析
args = parser.parse_args()
# 生成导出文件列表
file_list=[]
if os.path.isfile(args.input):
    file_list.append(args.input)
elif os.path.isdir(args.input):
    for maindir, subdir, file_name_list in os.walk(args.input):
        for filename in file_name_list:
            apath = os.path.join(maindir, filename)
            if apath.split('.')[-1] in ['xlsx','xlsm','xls']:
                file_list.append(apath)
dump=None
def dump_function(dump_type):
    global dump
print(args.input)
# print(args.excelpath)
# # 读取文件
# sys.argv=['0','.\\test_doublekey.xlsx','.\\export','lua',1]

# export_format=sys.argv[1]
# excel_path=sys.argv[2]
# export_path=sys.argv[3]
# path_check(export_path)
# workbook=xlrd.open_workbook(excel_path)
# WORKBOOK_NAME=os.path.split(excel_path)[1].split('.')[0]
# # 导入输出文件格式模板
# EXPORT_LANG=importlib.import_module(export_format)
# # 导出水平表和垂直表
# book_vertical,book_horizon=get_sheet_list(workbook)
# [gen_export_file_horizon(workbook.sheet_by_name(book)) for book in book_horizon]
# [gen_export_file_vertical(workbook.sheet_by_name(book)) for book in book_vertical]
