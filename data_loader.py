import re
import xlrd
import os
import sqlite3
import data_processor
import data_checker

def _clear(attri:list,value):
    for i in range(0,len(attri)):
        attri[i]=value
    pass

class _Head:
    name = ''
    name_cn=''
    type = ''
    count = 1
    stage = ''
    default = ''
    def __init__(self,head_data_raw):
        self.name_cn=head_data_raw['name_cn']
        if head_data_raw['name']!='':
            part = re.findall(r'\[([a-zA-Z]+)([0-9]*)\](.+)',
                        head_data_raw['name'])[0]
            self.type, self.count, self.name = part[0], 1 if (
                n := part[1]) == '' else int(n), part[2]
            self.stage, self.default = head_data_raw['stage'], head_data_raw['default']
        else:
            self.type='remark'
            self.name=self.count=self.stage=self.default=''





class Work:
    input = ''
    version = ''
    VALID_INPUT_FORMAT = ['xlsx', 'xlsm', 'xls']
    book_list = {}
    localization:sqlite3.Connection

    class OutputPath:
        client = ''
        server = ''

    class OutputFormat:
        client = ''
        server = ''

    class DumpMethod:
        client = None
        server = None

    @staticmethod
    def update(args):
        #初始化导出配置
        Work.input =path = args.input
        Work.version = args.version
        Work.OutputPath.client = args.output_client
        Work.OutputPath.server = args.output_server
        Work.OutputFormat.client = args.format_client
        Work.OutputFormat.server = args.format_server
        # 初始化要导出的文件列表
        if os.path.isfile(path) and path.split('.')[-1] in Work.VALID_INPUT_FORMAT:
            name=os.path.split(path)[1].split('.')[0]
            Work.book_list[name]=path
        elif os.path.isdir(path):
            for maindir, subdir, file_name_list in os.walk(path):
                for _file in file_name_list:
                    full_path = os.path.join(maindir, _file)
                    if _file.split('.')[-1] in Work.VALID_INPUT_FORMAT:
                        Work.book_list[_file.split('.')[0]]=full_path
        # 初始化本地化内容
        Work.localization = sqlite3.connect(os.path.realpath(Work.OutputPath.client + "..\Localization.db"))
        Work.localization.execute("""create table if not exists Localization (key text primary key,version text,origin text,zh_CN text)""")
        Work.localization.commit()
        #初始化导出格式
        def set_format(format_type):
            if format_type == 'lua':
                from format import lua
                return lua.dumps
            elif format_type == 'yaml':
                import yaml
                return yaml.dump
            elif format_type == 'json':
                import json
                return json.dumps
        Work.DumpMethod.client=set_format(Work.OutputFormat.client)
        Work.DumpMethod.server=set_format(Work.OutputFormat.server)

class Workbook:
    obj=None
    name = ''
    path = ''
    sheet_list=[]
    @staticmethod
    def update(name):
        #清空原始值
        Workbook.sheet_list=[]
        #赋新值
        Workbook.obj=xlrd.open_workbook(Work.book_list[name])
        Workbook.name=name
        Workbook.path=Work.book_list[Workbook.name]
        for sheet in Workbook.obj.sheets():
            if sheet.name[0]!="_":
                Workbook.sheet_list.append(sheet.name)


class Sheet:
    obj = None
    name = ''
    type = ''
    data_start_row=0
    row_max = 0
    col_max = 0
    heads = []
    master_key_col = -1
    sub_key_col = -1
    sub_key_name=''
    data_client = {}
    data_server = {}        
    data_range=[[0,0],[0,0]] # 开始、结束单元格的xlrd地址

    @staticmethod
    def update(sheet:xlrd.sheet):
        # 清空基础信息
        Sheet.data_client=Sheet.data_server={}
        Sheet.master_key_col=Sheet.sub_key_col=-1
        # 更新sheet信息
        Sheet.obj=sheet
        Sheet.row_max=sheet.nrows
        Sheet.col_max=sheet.ncols
        Sheet.heads=[]
        # 精简型配置表
        if sheet.name[0]=='@':
            Sheet.name= sheet.name[1:]
            Sheet.type='Lite'
            Sheet.data_start_row=1
            # 更新表头信息
            for i in range(1,Sheet.row_max):
                head_data_raw={
                    'name_cn': Sheet.obj.cell(i,0).value,
                    'name': Sheet.obj.cell(i,1).value,
                    'stage': Sheet.obj.cell(i,2).value,
                    'default': Sheet.obj.cell(i,3).value}
                data_checker.check_head(head_data_raw) #检查表头数据原值
                Sheet.heads.append(_Head(head_data_raw))
            Sheet.data_range=[[Sheet.data_start_row,3],[Sheet.row_max,4]]
        # 普通配置表
        else:
            Sheet.name= sheet.name
            Sheet.type='Normal_SingleKey'
            Sheet.data_start_row=4
            # 更新表头信息
            for i in range(0,Sheet.col_max):
                head_data_raw={
                    'name_cn': Sheet.obj.cell(0,i).value,
                    'name': Sheet.obj.cell(1,i).value,
                    'stage': Sheet.obj.cell(2,i).value,
                    'default': Sheet.obj.cell(3,i).value}
                data_checker.check_head(head_data_raw) #检查表头数据原值
                Sheet.heads.append(_Head(head_data_raw))
                if Sheet.heads[-1].type[1:]=='key':
                    if Sheet.master_key_col==-1:
                        Sheet.master_key_col=i
                    elif Sheet.sub_key_col==-1:
                        Sheet.sub_key_col=i
                        Sheet.sub_key_name=Sheet.heads[-1].name
                        Sheet.type='Normal_DoubleKey'
                    else:
                        data_checker.too_many_keys(head_data_raw['name_cn'])
            Sheet.data_range=[[Sheet.data_start_row,Sheet.master_key_col],[Sheet.row_max,Sheet.col_max]]
                    

class Cell:
    address_xlrd = [0, 0]
    address_excel=[-1,-1]
    value= ''
    value_raw= ''
    master_key = -1
    sub_key = -1
    dict_key = ''
    array_index = ''
    head:_Head
    @staticmethod
    def update(xlrd_row,xlrd_col):
        def update_dict(dict_ref:dict,keys:list,value):
            '''递归更新字典'''
            if keys.__len__()==1:
                dict_ref[keys[0]]=value
            else:
                if dict_ref.__contains__(keys[0]):
                    update_dict(dict_ref[keys[0]],keys[1:],value)
                else:
                    dict_ref[keys[0]]={}
                    update_dict(dict_ref[keys[0]],keys[1:],value)
        # 更新Cell地址
        Cell.address_xlrd=[xlrd_row,xlrd_col]
        Cell.address_excel=[xlrd_row+1,xlrd_col+1]
        # 更新head信息和原始值
        if Sheet.type=='Lite':
            Cell.head=Sheet.heads[xlrd_row-Sheet.data_start_row]
            Cell.value_raw=Cell.head.default
            Cell.master_key=Cell.head.name
            data_checker.check_data_type(Cell.value_raw,Cell.head._type) #检查值
            Cell.value=data_processor.process_type[Cell.head.type](Cell.value_raw)
            update_dict(Sheet.data_client,[Cell.master_key],Cell.value)
        else:
            Cell.head=Sheet.heads[xlrd_col]
            if Cell.head.type!='remark':
                Cell.value_raw=Sheet.obj.cell(xlrd_row,xlrd_col).value
                data_checker.check_data_type(Cell.value_raw,Cell.head._type)
                Cell.value=data_processor.process_type[Cell.head.type](Cell.value_raw) #检查值
                Cell.master_key=data_processor.process_type[Sheet.heads[Sheet.master_key_col].type](Sheet.obj.cell(xlrd_row, Sheet.master_key_col).value)
                if Sheet.type=='Normal_SingleKey':
                    Cell.sub_key=-1
                    update_dict(Sheet.data_client,[Cell.master_key,Cell.head.name],Cell.value)
                elif Sheet.type=='Normal_DoubleKey':
                    Cell.sub_key=data_processor.process_type[Sheet.heads[Sheet.sub_key_col].type](Sheet.obj.cell(xlrd_row, Sheet.sub_key_col).value)
                    if xlrd_col!=Sheet.master_key_col and xlrd_col!=Sheet.sub_key_col:
                        if xlrd_col<Sheet.sub_key_col:
                            update_dict(Sheet.data_client,[Cell.master_key,Cell.head.name],Cell.value)
                        else:
                            update_dict(Sheet.data_client,[Cell.master_key,Sheet.sub_key_name,Cell.sub_key,Cell.head.name],Cell.value)
                pass


def exception_str():
    return "【表名：%s；sheet名：%s；字段名：%s；主键：%s；子键：%s；行：%s；列：%s；】" % \
        (Workbook.name, Sheet.name, Cell.head.name,
         Cell.master_key, Cell.sub_key, Cell.address_excel[0], Cell.address_excel[1])


def text_key_str():
    key_str=''
    if Sheet.type=='Lite':
        key_str="%s.%s.%s" % \
            (Workbook.name, Sheet.name,
             Cell.head.name)
    elif Sheet.type=='Normal_SingleKey':
        key_str="%s.%s.%s.%s" % \
            (Workbook.name, Sheet.name, Cell.master_key,
             Cell.head.name)
    elif Sheet.type=='Normal_DoubleKey':
        key_str="%s.%s.%s_%s.%s" % \
            (Workbook.name, Sheet.name, Cell.master_key, Cell.sub_key,
             Cell.head.name)
    if Cell.head.type=='d':
        key_str +='.'+str(Cell.dict_key)
    if Cell.array_index != '':
        key_str+='.'+'_'.join(map(str, Cell.array_index))
    return key_str


def output_str(stage):
    if stage=='client':
        return Work.DumpMethod.client(Sheet.data_client)
    else:
        return Work.DumpMethod.server(Sheet.data_server)

def file_name(stage: str):
    if stage == 'client':
        path, format = Work.OutputPath.client, Work.OutputFormat.client
    elif stage == 'server':
        path, format = Work.OutputPath.server, Work.OutputFormat.server
    return '%s\\%s_%s.%s' % (path, Workbook.name, Sheet.name, format)
