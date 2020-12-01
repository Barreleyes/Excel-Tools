import re
import xlrd
import os
import sqlite3

def _clear(attri:list,value):
    for i in range(0,len(attri)):
        attri[i]=value
    pass

class _Field:
    name = ''
    type = ''
    count = 1
    stage = ''
    default = ''
    col = 0

    def __init__(self, data: [], head_index, col_index):
        part = re.findall(r'\[([a-zA-Z]+)([0-9]*)\](.+)',
                          data[head_index['name']])[0]
        self.type, self.count, self.name = part[0], 1 if (
            n := part[1]) == '' else int(n), part[2]
        self.stage, self.default = data[head_index['stage']
                                        ], data[head_index['default']]
        self.col = col_index


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
    def basic_info(args):
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
        Work.localization.execute("""create table if not exists Localization (
            key text primary key,
            version text,
            zh_CN text)""")
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
    sheet_list_normal = []
    sheet_list_lite = []
    @staticmethod
    def basic_info(name):
        _clear([Workbook.name,Workbook.path],'')
        _clear([Workbook.sheet_list_lite,Workbook.sheet_list_normal],[])
        Workbook.obj=xlrd.open_workbook(Work.book_list[name])
        Workbook.name=name
        Workbook.path=Work.book_list[Workbook.name]
        for sheet in Workbook.obj.sheets():
            if sheet.name[0] == '@':
                Workbook.sheet_list_lite.append(sheet.name)
            elif sheet.name[0] != '_':
                Workbook.sheet_list_normal.append(sheet.name)


class Sheet:
    obj = None
    name = ''
    type = ''
    data_row=0
    row_max = 0
    col_max = 0
    head_index={}
    head = []
    mainkey_col = -1
    subkey_col = -1
    mainkey_list = []
    subkey_list = []
    data_group = {}
    data_client = {}
    data_server = {}
    @staticmethod
    def add_head(data: [], col_index):
        if data[Sheet.head_index['name']]!='':
            Sheet.head.append(c_field:=_Field(data,Sheet.head_index,col_index))
            if c_field.type[1:]=='key':
                if Sheet.mainkey_col== -1:
                    Sheet.mainkey_col= c_field.col
                else:
                    Sheet.subkey_col = c_field.col
        else:
            Sheet.head.append('')

    @staticmethod
    def basic_info(sheet:xlrd.sheet):
        Sheet.obj=sheet
        Sheet.row_max=sheet.nrows
        Sheet.col_max=sheet.ncols
        normal = {'name_cn': 0, 'name': 1, 'stage': 2, 'default': 3}
        lite = {'name_cn': 0, 'name': 1, 'stage': 2, 'default': 3}
        if sheet.name[0]=='@':
            Sheet.name= sheet.name[1:]
            Sheet.type='Lite'
            Sheet.data_row=2
            Sheet.head_index=lite
        else:
            Sheet.name= sheet.name
            Sheet.type='Normal'
            Sheet.data_row=4
            Sheet.head_index=normal

    @staticmethod
    def clear():
        _clear([Sheet.name,Sheet.type],'')
        _clear([Sheet.head_index,Sheet.data_group,Sheet.data_client,Sheet.data_server],{})
        _clear([Sheet.head,Sheet.mainkey_list,Sheet.subkey_list],[])
        _clear([Sheet.obj,Sheet.mainkey_col,Sheet.subkey_col],-1)





class Cell:
    address = [0, 0]
    value= ''
    orig_value=''
    mainkey = ''
    subkey = ''
    dict_key = ''
    array_index = [0,0]
    field:_Field

    @staticmethod
    @property
    def row():
        return Cell.address[0]+1

    @staticmethod
    @property
    def col():
        return Cell.address[1]+1

    @staticmethod
    def clear():
        _clear([Cell.mainkey,Cell.subkey,Cell.dict_key,Cell.array_index],'')



def exception_str():
    return "【表名：%s；sheet名：%s；字段名：%s；主键：%s；子键：%s；行：%s；列：%s；】" % \
        (Workbook.name, Sheet.name, Cell.field.name,
            Cell.mainkey, Cell.subkey, Cell.row, Cell.col)


def text_key_str():
    if Cell.array_index == '':
        return "%s_%s_%s_%s_%s" % \
            (Workbook.name, Sheet.name,
                Cell.field.name, Cell.mainkey, Cell.subkey)
    else:
        array_index=''.join(map(str, Cell.array_index)) if Cell.array_index!=[0,0] else ''
        return ("%s_%s_%s_%s_%s_%s_%s" %
                (Workbook.name, Sheet.name,
                    Cell.field.name, Cell.mainkey, Cell.subkey, Cell.dict_key, array_index)).replace('__', '_')

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
