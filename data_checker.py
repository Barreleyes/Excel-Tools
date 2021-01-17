import re
import data_processor
import data_loader
EXCEPTION_COLLECTOR = []
EXCEPTION_TYPE_TEXT={
    '101':"表头错误[extra_text]：数组字段必须填写下标",
    '102':"表头错误[extra_text]：数组字段下标必须为不为0的阿拉伯数字",
    '103':"表头错误[extra_text]：填写了不支持的字段类型。",
    '104':"表头错误[extra_text]：字段类型必须包裹在[]中",
    '105':"表头错误[extra_text]：未填写字段名",
    '106':"表头错误[extra_text]：字段名必须以字母打头，且仅包含字母、数字或下划线",
    '107':"表头错误[extra_text]：配置表中至多有2个键",
    '111':"导出平台类型错误[extra_text]：仅可使用“all”,“client”,“server”三种类型",
    '121':"默认值错误[extra_text]：默认值与字段类型不符",
    '131':"字段中文名错误：未填写字段中文名",
    '201':"填写了重复的主键",
    '202':"填写了重复的子键",


}

def too_many_keys(name_cn):
    update_exception(107,name_cn)

def update_exception(exception_type,extra_text=''):
    EXCEPTION_COLLECTOR.append({
        'workbook':data_loader.Workbook.name,
        'sheet':data_loader.Workbook.name,
        'row':data_loader.Cell.address_excel[0],
        'col':data_loader.Cell.address_excel[1],
        'text':EXCEPTION_TYPE_TEXT[exception_type].replace('extra_text',extra_text)
    })


def check_head(head_data_raw):
    if head_data_raw['name_cn']=='':
        update_exception('131')
    parse=re.findall(r'\[([a-zA-Z]+)([0-9]*)\](.+)',head_data_raw['name'])[0]
    if parse[0]=='':
        update_exception('104',head_data_raw['name_cn'])
    elif parse[0] not in data_processor.process_type:
        update_exception('103',head_data_raw['name_cn'])
    if parse[3]=='':
        update_exception('105',head_data_raw['name_cn'])
    elif (not parse[3].isascii()) or (not parse[3][0].isalpha()):
        update_exception('106',head_data_raw['name_cn'])
    if not head_data_raw['stage'] in ['all','client','server']:
        update_exception('111',head_data_raw['name_cn'])
    if not check_data_type(head_data_raw['default'],parse[0]):
        update_exception('121',head_data_raw['name_cn'])
    
def check_data_type(data_raw:str,_type):
    '''
    检查数据类型
    '''
    def str_contain(data,checkstr):
        r=True
        for i in checkstr:
            if i not in data:
                r=False
                break
        return r
    def check_arrayint(data):
        r=True
        for i in re.split('[;|]',data):
            if not i.isdigit():
                r=False
                break
        return r

    data_type = {
        "i": data_raw.isdigit(),
        "f": data_raw.isdigit(),
        "n": data_raw.isdigit(),
        "s": True,
        "t": True,
        "ai": check_arrayint(data_raw),
        "af": check_arrayint(data_raw),
        "an": check_arrayint(data_raw),
        "as": str_contain(data_raw,[';','|']),
        "at": str_contain(data_raw,[';','|']),
        "d": str_contain(data_raw,['#','=']),
        "ikey": data_raw.isdigit(),
        "skey": not data_raw.isdigit(),
    }
    return data_type[_type]

def check_master_key(key):
    if key in data_loader.Sheet.data_client.keys():
        update_exception('201')
def check_sub_key(key):
    if key in data_loader.Sheet.data_client[data_loader.Sheet.sub_key_name].keys():
        update_exception('202')
