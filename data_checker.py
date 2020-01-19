import re
def __excel_address_convert(old_address):
    '''
    将xlrd地址转换为excel地址
    '''
    return [x+1 for x in old_address]
def field_name(cell_address,field_type_range,data):
    '''
    cell_address:单元格地址[r,c]
    field_type_range:字段类型取值范围
    data：待检查的数据
    检查字段名中所包含的字段类型是否合法
    '''
    field_type=data.split('_')[0]
    if field_type[0]=='a':
        try:
            element_count=field_type.split('[')[1].split(']')[0]
        except IndexError:
            notice="字段类型错误[%d行,%d列]：数组字段必须填写下标"%(cell_address[0],cell_address[1])
            raise Exception(notice,cell_address)
        try:
            int(element_count)
        except ValueError:
            notice="字段类型错误[%d行,%d列]：数组字段下标必须为不为0的阿拉伯数字"%(cell_address[0],cell_address[1])
            raise Exception(notice,cell_address)
    if data.split('_')[0].split('[')[0] not in field_type_range:
        cell_address=__excel_address_convert(cell_address)
        notice="字段类型错误[%d行,%d列]：填写了不存在的字段类型。"%(cell_address[0],cell_address[1])
        raise Exception(notice,cell_address)
def field_range(cell_address,data):
    '''
    cell_address:单元格地址[r,c]
    data：待检查的数据
    检查配置表字段使用类型(前端、后端)及主、副键
    '''
    key_type=['key','subkey']
    use_range=['all','client','server']
    cell_address=__excel_address_convert(cell_address)
    split_Data = data.split('_')
    if split_Data[0] not in use_range:
        notice="字段使用范围错误[%d行,%d列]：仅可填写[all,server,client]。"%(cell_address[0],cell_address[1])
        raise Exception(notice,cell_address)
    if len(split_Data) !=1 and split_Data[1] not in key_type:
        notice="字段key类型错误[%d行,%d列]：仅可填写[mainkey,subkey]。"%(cell_address[0],cell_address[1])
        raise Exception(notice,cell_address)
def default_value(cell_address,field_type_range,data):
    
    pass