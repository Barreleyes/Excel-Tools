# --coding:utf-8--
'''
此模块用于将excel sheet转换为python字典
'''
import xlrd
import data_process
import data_checker
import current

sheet = current.Sheet


def cell(row, col):
    current.Cell.address = [row, col]
    current.Cell.orig_value=sheet.obj.cell(row, col).value
    return sheet.obj.cell(row, col).value


def load_sheet_normal():
    '''
    读取普通配置表
    '''

    # 读取表头信息
    def load_head_info():
        for col in range(0, sheet.col_max):
            head_data = []
            for row in range(0, sheet.data_row):
                head_data.append(cell(row, col))
            sheet.add_head(head_data, col)

    # 读取双建表分组
    def load_group_info():

        key_method = data_process.process_type[sheet.head[sheet.mainkey_col].type]
        for row in range(sheet.data_row, sheet.row_max):
            current.Cell.field = sheet.head[sheet.mainkey_col]
            mainkey = key_method(cell(row, sheet.mainkey_col))
            if mainkey not in sheet.data_group.keys():
                sheet.data_group[mainkey] = [row]
            else:
                sheet.data_group[mainkey].append(row)

    # 读取指定行字段区间内的数据，并储存在字典中
    def deserialize(row, col_start, col_end):
        client = {}
        server = {}
        for col in range(col_start, col_end):
            if  field:= sheet.head[col]:
                # 对excel表中的值进行处理，若值为空，则使用默认值
                current.Cell.clear()
                current.Cell.field=field
                value=current.Cell.value = data_process.process_type[current.Cell.field.type](
                    value if (value := cell(row, col)) else current.Cell.field.default)
                name,stage=current.Cell.field.name,current.Cell.field.stage
                if stage == 'all':
                    client[name] = value
                    server[name] = value
                elif stage == 'client':
                    client[name] = value
                else:
                    server[name] = value
        return client, server

    # 读取双建表
    def doublekey():
        for mainkey, rowlist in sheet.data_group.items():
            # 生成主键数据
            current.Cell.clear()
            current.Cell.mainkey = mainkey
            current.Cell.subkey = 'null'
            sheet.data_client[mainkey], sheet.data_server[mainkey] = deserialize(
                rowlist[0], sheet.mainkey_col + 1, sheet.subkey_col)
            # 生成子键数据
            subkey_item_client = {}
            subkey_item_server = {}
            subkey_name = sheet.head[sheet.subkey_col].name
            for sub_item_row in rowlist:
                current.Cell.field=sheet.head[sheet.subkey_col]
                subkey = current.Cell.subkey = subkey_method(
                    cell(sub_item_row, sheet.subkey_col))    
                sheet.subkey_list.append(subkey)
                subkey_item_client[subkey], subkey_item_server[subkey] = deserialize(
                    sub_item_row, sheet.subkey_col, sheet.col_max)
            sheet.data_client[mainkey][subkey_name], sheet.data_server[
                mainkey][subkey_name] = subkey_item_client, subkey_item_server
            sheet.subkey_list = []

    # 读取单键表
    def singlekey():
        current.Cell.subkey = 'null'
        sheet.mainkey_list = []
        for row in range(sheet.data_row, sheet.row_max):
            current.Cell.field=sheet.head[sheet.mainkey_col]
            mainkey = current.Cell.mainkey = mainkey_method(cell(row, sheet.mainkey_col))
            sheet.mainkey_list.append(mainkey)
            sheet.data_client[mainkey], sheet.data_server[mainkey] = deserialize(
                row, sheet.mainkey_col + 1, sheet.col_max)

    # 开始读取配置信息
    load_head_info()
    mainkey_method = data_process.process_type[sheet.head[sheet.mainkey_col].type]
    subkey_method = data_process.process_type[sheet.head[sheet.subkey_col].type]
    if sheet.subkey_col!=-1:
        load_group_info()
        doublekey()
    else:
        singlekey()


def load_sheet_lite():
    '''
    读取简化配置表
    '''

    current.Cell.mainkey = ''
    current.Cell.subkey = 'null'
    for row in range(sheet.data_row, sheet.row_max):
        head_data = []
        for col in range(0, 4):
            head_data.append(cell(row, col))
        current.Cell.field=current._Field(head_data,current.Sheet.head_index,row)
        sheet.mainkey_list.append(current.Cell.field.name)
        name,stage=current.Cell.field.name,current.Cell.field.stage
        value=data_process.process(current.Cell.field.type, current.Cell.field.default)
        if stage == 'all':
            sheet.data_client[name] = sheet.data_server[
                name] =value
        elif stage == 'client':
            sheet.data_client[name] = value
        elif stage=='server':
            sheet.data_server[name] = value


def load_sheet(sheet_obj: xlrd.sheet):
    sheet.basic_info(sheet_obj)
    current.Sheet.clear()
    current.Cell.clear()
    if sheet.type == 'Lite':
        load_sheet_lite()
    elif sheet.type == 'Normal':
        load_sheet_normal()
    sheet.data_client['TableVersion'] = sheet.data_server['TableVersion'] = current.Work.version