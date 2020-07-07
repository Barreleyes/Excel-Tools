import sys
import os
import argparse
import xlrd
import sheet_load
# 创建外部参数解析器
parser = argparse.ArgumentParser()
parser.add_argument('-i', '--input', default='.\\test_doublekey.xlsx',
                    help="excel文件或文件夹路径，默认读取【工具目录】")
parser.add_argument('-oc', '--output-client', default='.\\export\\client\\',
                    help="前端导出文件路径，默认在【工具目录\\export\\client】文件夹")
parser.add_argument('-os', '--output-server',default='.\\export\\server\\',
                    help="后端导出文件路径，默认在【工具目录\\export\\server】文件夹")
parser.add_argument('-fc', '--format-client', default='lua',
                    help="前端导出格式，可根据项目定制，默认为lua")
parser.add_argument('-fs', '--format-server', default='json',
                    help="后端导出格式，可根据项目定制，默认为lua")
args = parser.parse_args()

# 定义导出格式
format_client=None
format_server=None

def path_check(path):
    '''
    检查目标路径是否存在，若不存在则自动创建
    '''
    if not os.path.exists(path):
        os.makedirs(path)
        
def set_export_format():
    global format_client,format_server
    def set_format(format_type):
        if format_type=='lua':
            import lua
            return lua.dumps
        elif format_type=='yaml':
            import yaml
            return yaml.dump
        elif format_type=='json':
            import json
            return json.dumps
            pass
    format_client=set_format(args.format_client)
    format_server=set_format(args.format_server)
set_export_format()
# 生成导出文件列表
def get_sheet_list(workbook):
    '''
    :param workbook:xlrd的workbook对象
    
    :returns:(horizon,vertical)水平配置表和垂直配置表的字典列表，key是要导出的文件名，value是包含的sheet名
    '''
    horizon={}
    vertical={}
    def dict_append(dst_dict,name):
        if (n:=name.split('_')[0].replace('@','')) in dst_dict:
            dst_dict[n].append(name)
        else:
            dst_dict[n]=[]
            dst_dict[n].append(name)
    for sheet in workbook.sheets():
        sheet_name=sheet.name
        if sheet_name[0]=='@':
            dict_append(vertical,sheet_name)
        elif sheet.name[0]=='_':
            pass
        else:
            dict_append(horizon,sheet_name)
    return horizon,vertical
def get_file_list(path):
    file_list=[]
    if os.path.isfile(path):
        file_list.append(path)
    elif os.path.isdir(path):
        for maindir, subdir, file_name_list in os.walk(path):
            for filename in file_name_list:
                apath = os.path.join(maindir, filename)
                if apath.split('.')[-1] in ['xlsx','xlsm','xls']:
                    file_list.append(apath)
    return file_list

def save_file(data,path):
    f=open(path,'w',encoding='utf-8')
    f.write(data)
    f.close
# 导出配置表
def export_sheets(sheet_group,sheet_type):
    for k,v in sheet_group.items():
        key_list=[]
        data_client={}
        data_server={}
        for sheet in v:
            print("正在导出[%s]的[%s]..."%(workbook_name,sheet))
            sheet_data=sheet_load.SheetData(workbook_obj.sheet_by_name(sheet),sheet_type)
            sheet_data.mainkey_list=key_list
            data_client.update(sheet_data.data_client)
            data_server.update(sheet_data.data_server)
            key_list+=sheet_data.mainkey_list
        data_client=format_client(data_client)
        data_server=format_server(data_server)
        save_file(data_client,'%s\\%s_%s.%s'%(args.output_client,workbook_name,k,args.format_client))
        save_file(data_server,'%s\\%s_%s.%s'%(args.output_server,workbook_name,k,args.format_server))
for workbook in get_file_list(args.input):
    workbook_name=os.path.split(workbook)[1]
    workbook_obj=xlrd.open_workbook(workbook)
    sheet_horizon,sheet_vertical=get_sheet_list(workbook_obj)
    export_sheets(sheet_horizon,sheet_load.SheetType.HORIZON)
    export_sheets(sheet_vertical,sheet_load.SheetType.VERTICAL)
