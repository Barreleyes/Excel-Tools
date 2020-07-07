import xlrd
import re
import data_process
import data_checker
from enum import Enum
class SheetType(Enum):
    HORIZON=1
    VERTICAL=2
class SheetData:
    def __init__(self, sheet, sheet_type):
        # excel水平表头定义
        self.head_attribute_h = ['name_cn', 'name',
                                 'stage', 'default'] 
        # excel竖直表头定义 
        self.head_attribute_v = ['name_cn', 'name',
                                 'stage', 'value']  
        # excel竖直表读取起始行
        self.start_row_v = 2
        self.sheet = sheet
        self.head_info = []
        self.group_info = []
        self.mainkey_col = None
        self.subkey_col = None
        self.column_max = sheet.ncols
        self.row_max = sheet.nrows
        self.data_client = {}
        self.data_server = {}
        self.mainkey_list=[]
        self.__subkey_list=[]
        self.__current_cell = (0, 0)
        if sheet_type == SheetType.HORIZON:
            self.__load_head_info()
            self.__load_group_info()
            self.__load_sheet_horizon()
        elif sheet_type == SheetType.VERTICAL:
            self.__load_sheet_vertical()

    def cells(self, row, col):
        '''
        :param row:行号
        :param col:列号
        读取excel单元格数据
        '''
        return self.sheet.cell(row, col).value

    def __set_current_cell(self, row, col):
        '''
        储存当前读取的单元格地址，用于错误检查输出信息
        '''
        self.__current_cell = (row+1, col+1)

    def __field_name_analyse(self, field_name):
        '''字段名解析，返回字段类型，数组二维长度，字段名'''
        # data_checker.field_name(self.__current_cell,field_name)
        part = re.findall(r'\[([a-zA-Z]+)([0-9]*)\](.+)', field_name)[0]
        return part[0], 1 if (n:=part[1]) == '' else int(n), part[2]
    def __check_repeat(self,key,key_list):
        if key in key_list:
            data_checker.RaiseException.key_repeat(self.__current_cell)
        else:
            key_list.append(key)
    def __load_head_info(self):
        '''
        获取并分析水平表的表头信息。
        '''
        def head_info(col):
            info = {}
            for row in range(0, len(self.head_attribute_h)):
                self.__set_current_cell(row, col)
                row_tag = self.head_attribute_h[row]
                value = self.cells(row, col)
                if row_tag == 'name':
                    if value == '':
                        return info
                    info['type'], info['count'], info['name'] = self.__field_name_analyse(
                        value)
                    if 'key' in info['type']:
                        if self.mainkey_col and self.subkey_col:
                            data_checker.RaiseException.key_too_many(
                                self.__current_cell)
                        elif self.mainkey_col == None:
                            self.mainkey_col = col
                        else:
                            self.subkey_col = col
                elif row_tag == 'stage':
                    info['stage'] = value
                elif row_tag == 'default':
                    info['default'] = value
            return info
        for col in range(0, self.column_max):
            self.head_info.append(head_info(col))

    def __load_group_info(self):
        '''
        获取配置中的主副键组,存入二维数组
        '''
        if self.subkey_col:
            temp = []
            for row in range(len(self.head_attribute_h)-1, self.row_max):
                if (n:=self.cells(row, self.mainkey_col)) != '' and n != self.cells(row-1, self.mainkey_col):
                    temp.append(row)
            for i in range(0, len(temp)-1):
                self.group_info.append([temp[i], temp[i+1]-1])
            self.group_info.append([temp[-1], self.row_max-1])

    def head_info_check(self):
        if self.mainkey_col == None:
            data_checker.RaiseException.key_too_less(self.__current_cell)
        if self.head_info[self.mainkey_col]['stage'] == '':
            pass

    def __load_sheet_horizon(self):
        '''
        读取水平配置表，将结果保存在data_client和data_server中
        '''
        mainkey_type = self.head_info[self.mainkey_col]['type']
        def deserialize(row, col_start, col_end):
            '''
            读取指定行字段区间内的数据，并储存在字典中
            '''
            client = {}
            server = {}
            for col in range(col_start, col_end):
                if self.head_info[col]:
                    field_name, field_stage = self.head_info[col]['name'], self.head_info[col][
                        'stage']
                    # 对excel表中的值进行处理，若值为空，则使用默认值
                    field_value=data_process.process_type[self.head_info[col]['type']](value if (value := self.cells(row, col)) else self.head_info[col]['default'] , self.head_info[col]['count'])
                    self.__set_current_cell(row, col)
                    if field_stage == 'all':
                        client[field_name] = field_value
                    elif field_stage == 'client':
                        client[field_name] = field_value
                    else:
                        server[field_name] = field_value
            return client, server
        # 读取双建表
        if self.subkey_col:
            subkey_type = self.head_info[self.subkey_col]['type']
            for item in self.group_info:
                # 生成主键数据
                self.__set_current_cell(item[0], self.mainkey_col)
                mainkey = data_process.process_type[mainkey_type](
                    self.cells(item[0], self.mainkey_col))
                self.__check_repeat(mainkey,self.mainkey_list)
                self.data_client[mainkey], self.data_server[mainkey] = deserialize(
                    item[0], self.mainkey_col+1, self.subkey_col)
                # 生成子键数据
                subkey_item_client = {}
                subkey_item_server = {}
                subkey_name = self.head_info[self.subkey_col]['name']
                for item_sub in range(item[0], item[1]+1):
                    self.__set_current_cell(item_sub, self.subkey_col)
                    subkey = data_process.process_type[subkey_type](
                        self.cells(item_sub, self.subkey_col))
                    self.__check_repeat(subkey,self.__subkey_list)
                    subkey_item_client[subkey], subkey_item_server[subkey] = deserialize(
                        item_sub, self.subkey_col, self.column_max)
                self.data_client[mainkey][subkey_name], self.data_server[mainkey][subkey_name] = subkey_item_client, subkey_item_server
                self.__subkey_list=[]
        # 读取单键表
        else:
            for item in range(len(self.head_attribute_h)-1, self.row_max):
                self.__set_current_cell(item, self.mainkey_col)
                mainkey = data_process.process_type[mainkey_type](
                    self.cells(item, self.mainkey_col))
                # 检查主键重复
                self.__check_repeat(mainkey,self.mainkey_list)
                if mainkey in self.mainkey_list:
                    data_checker.RaiseException.key_repeat(self.__current_cell)
                self.mainkey_list.append(mainkey)
                
                self.data_client[mainkey], self.data_server[mainkey] = deserialize(
                    item, self.mainkey_col+1, self.column_max)

    def __load_sheet_vertical(self):
        '''
        读取垂直配置表
        '''
        col_name = self.head_attribute_v.index('name')
        col_stage = self.head_attribute_v.index('stage')
        col_value = self.head_attribute_v.index('value')
        for row in range(self.start_row_v-1, self.row_max):
            self.__set_current_cell(row,col_name)
            field_type, filed_count, field_name = self.__field_name_analyse(
                self.cells(row, col_name))
            self.__check_repeat(field_name,self.mainkey_list)
            field_value = data_process.process_type[field_type](
                self.cells(row, col_value),filed_count)
            field_stage = self.cells(row, col_stage)
            if field_stage == 'all':
                self.data_client[field_name] = self.data_server[field_name] = field_value
            elif field_stage == 'client':
                self.data_client[field_name] = field_value
            else:
                self.data_server[field_name] = field_value


# a = xlrd.open_workbook('.\\test_doublekey.xlsx').sheet_by_name('test')
# b = SheetData(a, SheetType.HORIZON)
# print('server' in b.head_info)
# c = 1
