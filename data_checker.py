import re
from data_processor import process_type as field_type_range
import data_loader
exception_collector = []

def field_name():
    '''
    cell_address:单元格地址[r,c]
    data：待检查的数据
    检查字段名中所包含的字段类型是否合法
    '''
    field_type = data.split('_')[0]
    if field_type[0] == 'a':
        try:
            element_count = field_type.split('[')[1].split(']')[0]
        except IndexError:
            raise Exception(cell_address_format + "\n字段类型错误：数组字段必须填写下标")
        try:
            int(element_count)
        except ValueError:
            raise Exception(cell_address_format +
                            "\n字段类型错误：数组字段下标必须为不为0的阿拉伯数字")
    if data.split('_')[0].split('[')[0] not in field_type_range:
        raise Exception(cell_address_format + "\n字段类型错误：填写了不支持的字段类型。")


def default_value(cell_address, field_type_range, data):

    pass


class RaiseException:
    @staticmethod
    def key_too_many():
        raise Exception(cell_address_format() + "\n键错误：单张配置表最多有两个键")

    @staticmethod
    def key_too_less():
        raise Exception(cell_address_format() + "\n键错误：单张配置表至少要有一个键")

    @staticmethod
    def stage_error():
        raise Exception(cell_address_format() +
                        "\n字段使用范围错误[%d行,%d列]：仅可填写[all,server,client]。")

    @staticmethod
    def key_repeat():
        raise Exception(cell_address_format() + "\n键错误[%d行,%d列]：填写了重复的键。")
