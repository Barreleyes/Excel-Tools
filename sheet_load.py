import xlrd
import re
import data_process


class SheetData:
    __sheet_start_row = 4     # 水平表起始数据行
    __sheet_h_name_cn = 0     # 水平表表头-中文名
    __sheet_h_name_en = 1     # 水平表表头-英文名&字段类型
    __sheet_h_range = 2       # 水平表表头-使用范围
    __sheet_h_default = 3     # 水平表表头-默认值
    __sheet_v_start_row = 0   # 垂直表起始数据行
    __sheet_v_name_cn = 0     # 垂直表表头-中文名
    __sheet_v_name_en = 1     # 垂直表表头-英文名&字段类型
    __sheet_v_value = 2       # 垂直表表头-值

    def __init__(self, sheet, sheet_type):
        self.head_info = {}
        self.group_info = []
        self.mainkey_col = 0
        self.subkey_col = 0
        self.key_count = 0
        self.__load_head_info(sheet)
        self.data = self.__load_sheet_horizon(
            sheet) if sheet_type == 0 else self.__load_sheet_vertical(sheet)

    def __field_name_analyse(self, field_name):
        part = re.search(r'\[(.+)\](.+)', field_name)
        return part.group(2), part.group(1)

    def __load_head_info(self, sheet):
        '''
        :param sheet:sheet对象
        获取并分析表头信息，然后储存在字典中。
        '''
        column_max = sheet.ncols
        for i in range(0, column_max):
            self.head_info[i] = {}
            field_name_cell = sheet.cell(self.__sheet_h_name_en, i).value
            field_range_cell = sheet.cell(self.__sheet_h_range, i).value
            default_value_cell = sheet.cell(self.__sheet_h_default, i).value
            if field_name_cell != '':
                # 表头信息整理
                field_name, field_type = self.__field_name_analyse(
                    field_name_cell)
                self.head_info[i]['name'] = field_name
                self.head_info[i]['type'] = field_type
                self.head_info[i]['default'] = default_value_cell
                self.head_info[i]['platform'] = field_range_cell
                if self.head_info[i]['type'][0] == 'a':
                    self.head_info[i]['count'] = int(
                        self.head_info[i]['type'][-1])
                    self.head_info[i]['type'] = field_type[:-1]
                else:
                    self.head_info[i]['count'] = 0
                    pass
                if 'key' in self.head_info[i]['type']:
                    self.key_count += 1
                    if self.key_count == 1:
                        self.mainkey_col = i
                    elif self.key_count == 2:
                        self.subkey_col = i
                        self.head_info[i]['type'] += '_sub'
            else:
                continue
        if self.key_count == 0:
            raise Exception("单张配置表应至少有一个主键")
        elif self.key_count == 2:
            start = self.__sheet_start_row
            for row in range(5, sheet.nrows):
                if sheet.cell(row, 0).value != '' or row == sheet.nrows-1:
                    if row == sheet.nrows-1:
                        self.group_info.append([start, row])
                    else:
                        self.group_info.append([start, row-1])
                        start = row
        elif self.key_count > 2:
            raise Exception("单张配置表最多有两个键")

    def __load_sheet_horizon(self, sheet):
        '''
        :param sheet:sheet对象
        将水平配置表sheet中的数据读取为python 字典
        '''
        temp_data = {}
        mainkey_type = self.head_info[self.mainkey_col]['type']

        def deserialize(row, col_start, col_end):
            data_temp = {}
            for col in range(col_start, col_end):
                if self.head_info[col]:
                    data_temp[self.head_info[col]['name']] = data_process.process_type[self.head_info[col]['type']](
                        sheet.cell(row, col).value)
            return data_temp
        if self.subkey_col == 0:
            for row in range(self.__sheet_start_row, sheet.nrows):
                key = data_process.process_type[mainkey_type](
                    sheet.cell(row, self.mainkey_col).value)
                temp_data[key] = deserialize(
                    row, self.mainkey_col+1, sheet.nrows)
            pass
        else:
            subkey_type = self.head_info[self.subkey_col]['type']
            for item in self.group_info:
                key = data_process.process_type[mainkey_type](
                    sheet.cell(item[0], self.mainkey_col).value)
                temp_data[key] = deserialize(
                    item[0], self.mainkey_col+1, self.subkey_col)
                subkey_item = {}
                for item_sub in range(item[0], item[1]+1):
                    subkey = data_process.process_type[subkey_type](
                        sheet.cell(item_sub, self.subkey_col).value)
                    subkey_item[subkey] = deserialize(
                        item_sub, self.subkey_col, sheet.ncols)
                temp_data[key][self.head_info[self.subkey_col]
                               ['name']] = subkey_item
                pass
        return temp_data

    def __load_sheet_vertical(self, sheet):
        temp_data = {}
        for row in range(self.__sheet_v_start_row, sheet.nrows):
            field_name, field_type = self.__field_name_analyse(
                sheet.cell(row, self.__sheet_v_name_en).value)
            temp_data[field_name] = data_process.process_type[field_type](
                sheet.cell(row, self.__sheet_v_value).value)
        return temp_data
