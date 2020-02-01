#--coding:utf-8--
'''
此模块用于将excel转换为python字典
'''
import importlib
import xlrd
import sys,os
import data_checker
import re
import data_process
SHEET_H_START_ROW=4 # 水平表起始数据行
SHEET_H_NAME_CN=0   # 水平表表头-中文名
SHEET_H_NAME_EN=1   # 水平表表头-英文名&字段类型
SHEET_H_RANGE=2     # 水平表表头-使用范围
SHEET_H_DEFAULT=3   # 水平表表头-默认值
SHEET_V_START_ROW=0 # 垂直表起始数据行
SHEET_V_NAME_CN=0   # 垂直表表头-中文名
SHEET_v_NAME_EN=1   # 垂直表表头-英文名&字段类型
SHEET_v_VALUE=2     # 垂直表表头-值
SHEET_HEAD_INFO=None
def field_name_analyse(field_name):
    part=re.search(r'\[(.+)\](.+)',field_name)
    return part.group(2),part.group(1)
def load_head_info(sheet):
    '''
    :param sheet:sheet对象
    获取并分析表头信息，然后储存在字典中。
    '''
    column_max=sheet.ncols
    head_info={}
    group_info=[]
    mainkey_col=0
    subkey_col=0
    key_count=0
    for i in range(0,column_max):
        head_info[i]={}
        field_name_cell=sheet.cell(SHEET_H_NAME_EN,i).value
        field_range_cell=sheet.cell(SHEET_H_RANGE,i).value
        default_value_cell=sheet.cell(SHEET_H_DEFAULT,i).value
        if field_name_cell =='':
            continue
        else:
            #表头内容检查
            # data_checker.field_name([1,i],field_type,field_name_cell)
            # data_checker.field_range([2,i],field_range_cell)
            #表头信息整理
            field_name,field_type=field_name_analyse(field_name_cell)
            head_info[i]['name']=field_name
            head_info[i]['type']=field_type
            head_info[i]['default']=default_value_cell
            head_info[i]['platform']=field_range_cell
            if head_info[i]['type'][0]=='a':
                head_info[i]['count']=int(head_info[i]['type'][-1])
                head_info[i]['type']=field_type[:-1]
            else:
                head_info[i]['count']=0
                pass
            if 'key' in head_info[i]['type']:
                key_count+=1
                if key_count==1:
                    mainkey_col=i
                elif key_count==2:
                    subkey_col=i
                    head_info[i]['type']+='_sub'
    if key_count==0:
        raise Exception("单张配置表应至少有一个主键")
    elif key_count==2:
        start=SHEET_H_START_ROW
        for row in range(5,sheet.nrows):
            if sheet.cell(row,0).value !='' or row == sheet.nrows-1:
                if row==sheet.nrows-1:
                    group_info.append([start,row])
                else:
                    group_info.append([start,row-1])
                    start=row
    elif key_count>2:
        raise Exception("单张配置表最多有两个键")
    global SHEET_HEAD_INFO
    SHEET_HEAD_INFO=head_info
    return head_info,mainkey_col,subkey_col,group_info
def load_sheet_horizon(sheet):
    '''
    :param sheet:sheet对象
    将水平配置表sheet中的数据读取为python 字典
    '''
    head_info,mainkey_col,subkey_col,group_info=load_head_info(sheet)
    temp_data={}
    mainkey_type=head_info[mainkey_col]['type']
    def deserialize(row,col_start,col_end):
        data_temp={}
        for col in range(col_start,col_end):
            if head_info[col]:
                data_temp[head_info[col]['name']]=data_process.process_type[head_info[col]['type']](sheet.cell(row,col).value)
        return data_temp
    if subkey_col==0:
        for row in range(SHEET_H_START_ROW,sheet.nrows):
            key=data_process.process_type[mainkey_type](sheet.cell(row,mainkey_col).value)
            temp_data[key]=deserialize(row,mainkey_col+1,sheet.nrows)
        pass
    else:
        subkey_type=head_info[subkey_col]['type']
        for item in group_info:
            key=data_process.process_type[mainkey_type](sheet.cell(item[0],mainkey_col).value)
            temp_data[key]=deserialize(item[0],mainkey_col+1,subkey_col)
            subkey_item={}
            for item_sub in range(item[0],item[1]+1):
                subkey=data_process.process_type[subkey_type](sheet.cell(item_sub,subkey_col).value)
                subkey_item[subkey]=deserialize(item_sub,subkey_col,sheet.ncols)
            temp_data[key][head_info[subkey_col]['name']]=subkey_item
            pass
    return temp_data
def load_sheet_vertical(sheet):
    temp_data={}
    for row in range(SHEET_V_START_ROW,sheet.nrows):
        field_name,field_type=field_name_analyse(sheet.cell(row,SHEET_v_NAME_EN).value)
        temp_data[field_name]=data_process.process_type[field_type](sheet.cell(row,SHEET_v_VALUE).value)
    return temp_data
# 读取文件
sys.argv=['0','lua','.\\test_doublekey.xlsx','.\\export']
sys.path.append('.\\templete')
export_format=sys.argv[1]
excel_path=sys.argv[2]
export_path=sys.argv[3]
workbook=xlrd.open_workbook(excel_path)
import sheet_load
a=sheet_load.SheetData(workbook.sheet_by_name("test"),0)
a=sheet_load.SheetData
a=1
# import yaml
# a=yaml.dump(a)
# open('.\\1.yaml','w').write(a)
# b=yaml.load(open('.\\1.yaml','r'))
# print(b)
# WORKBOOK_NAME=os.path.split(excel_path)[1].split('.')[0]
# import lua
# a=lua.dump(a)
# print(a)
# lua.dump_file(a,'.\\','test')
# open('.\\teslua.lua','w',encoding='utf-8').write(a)