import re
from data_process import process_type
field_type_range=process_type
def field_name(cell_address,data):
    '''
    cell_address:单元格地址[r,c]
    data：待检查的数据
    检查字段名中所包含的字段类型是否合法
    '''
    field_type=data.split('_')[0]
    if field_type[0]=='a':
        try:
            element_count=field_type.split('[')[1].split(']')[0]
        except IndexError:
            notice="字段类型错误[%d行,%d列]：数组字段必须填写下标"%cell_address
            raise Exception(notice,cell_address)
        try:
            int(element_count)
        except ValueError:
            notice="字段类型错误[%d行,%d列]：数组字段下标必须为不为0的阿拉伯数字"%cell_address
            raise Exception(notice,cell_address)
    if data.split('_')[0].split('[')[0] not in field_type_range:
        notice="字段类型错误[%d行,%d列]：填写了不存在的字段类型。"%cell_address
        raise Exception(notice,cell_address)

def default_value(cell_address,field_type_range,data):
    
    pass
class RaiseException:
    @staticmethod
    def key_too_many(cell_address):
        raise Exception("键错误[%d行,%d列]：单张配置表最多有两个键"%cell_address)
    @staticmethod
    def key_too_less(cell_address):
        raise Exception("键错误：单张配置表至少要有一个键")
    @staticmethod
    def stage_error(cell_address):
        raise Exception("字段使用范围错误[%d行,%d列]：仅可填写[all,server,client]。"%cell_address)
    @staticmethod
    def key_repeat(cell_address):
        raise Exception("键错误[%d行,%d列]：填写了重复的键。"%cell_address)
