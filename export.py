import importlib
import xlrd
import sys
import data_checker
import re
# 读取文件
sys.argv=['0','lua','.\\test.xlsx','.\\export']
sys.path.append('.\\templete')
export_format=sys.argv[1]
excel_path=sys.argv[2]
export_path=sys.argv[3]
workbook=xlrd.open_workbook(excel_path)
# 导入输出文件格式模板
export_lang=importlib.import_module(export_format)
# field_type=export_lang.field_type
# 获取excel配置表sheet列表，忽略说明，分割普通表和设置项。
def get_sheet_list(workbook):
    '''
    excel_path:excel文件路径
    根据类型将excel中的sheet进行分割
    '''
    vertical=[]
    horizon=[]
    for sheet in workbook.sheets():
        if sheet.name[0]=='@':
            vertical.append(sheet.name[1:])
        elif sheet.name[0]=='_':
            pass
        else:
            horizon.append(sheet.name)
    return vertical,horizon

#获取表头文件信息
def get_head_info(sheet):
    '''
    sheet:sheet对象
    获取并分析表头信息，然后储存在字典中。
    '''
    column_max=sheet.ncols
    head_info={}
    subkey_col=0
    if 'key' not in sheet.cell(2,0).value:
        raise Exception("主键错误，配置表A3只能填写[all_key,client_key,server_key]")
    for i in range(0,column_max):
        head_info[i]={}
        field_name_cell=sheet.cell(1,i).value
        field_range_cell=sheet.cell(2,i).value
        default_value_cell=sheet.cell(3,i).value
        if field_name_cell =='':
            continue
        else:
            #表头内容检查
            # data_checker.field_name([1,i],field_type,field_name_cell)
            # data_checker.field_range([2,i],field_range_cell)
            #表头信息整理
            part=re.search(r'\[(.+)\](.+)',field_name_cell)
            head_info[i]['name']=part.group(2)
            head_info[i]['type']=part.group(1)
            if head_info[i]['type'][0]=='a':
                head_info[i]['count']=int(head_info[i]['type'][-1])
                head_info[i]['type']=part.group(1)[:-1]
            else:
                head_info[i]['count']=0
                pass
            head_info[i]['default']=default_value_cell
            head_info[i]['platform']=field_range_cell.split('_')[0]
            if '_' in field_range_cell:
                head_info[i]['key']=field_range_cell.split('_')[1]
                if 'subkey' in field_range_cell:
                    subkey_col=i
            else:
                head_info[i]['key']=''
    return head_info,subkey_col
def get_key_info(sheet):
    '''
    sheet:sheet对象
    获取配置表单、双键结构，然后储存在数组中。
    '''
    
def gen_export_file(sheet):
    head_info,subkey_col=get_head_info(sheet)
    # 获取导出文件头尾
    text_head=export_lang.HEAD
    text_tail=export_lang.TAIL
    # 生成导出文件路径及文件名
    lua_name=export_path + '\\' + sheet.name + '.lua'
    # 生成导出文件
    def id_process(id):
        '''
        判断ID是整数还是字符串
        '''
        if isinstance(id,str):
            return '\"' + id +'\"'
        else:
            return int(id)
    export_file=open(lua_name,'w',encoding='utf-8')
    export_file.write(text_head)
    for row in range(4,sheet.nrows):       
        item_id=id_process(sheet.cell(row,0).value)
        item_data=''
        # 生成一个配置项的数据
        for column in range(1,sheet.ncols):
            if head_info[column]:
                item_data+=export_lang.value_format(head_info[column],sheet.cell(row,column).value)
        item_text=export_lang.ITEM.format(id=item_id,data=item_data)
        if item_id != '':
            export_file.write(item_text)
        pass
    export_file.write(text_tail)
    def item_process(id,row,col_start,col_end,head_info):
        '''
        :param id: 配置id
        :param row: 配置所在行
        :param col_start: 导出起始列
        :param col_end: 导出结束列
        :param head_info: 表头信息
        导出配置数据
        '''
        item_id= item_id= '\"' + id +'\"' if isinstance(id,str) else int(id)
        item_data=''
        for column in range(col_start,col_end):
            if head_info[column]:
                item_data+=export_lang.value_format(head_info[column],sheet.cell(row,column).value)
        item_data+='--subjoin--'
        return export_lang.ITEM.format(id=item_id,data=item_data)
b=get_sheet_list(workbook)
# c=get_head_info(workbook.sheet_by_name('test'))
gen_export_file(workbook.sheet_by_name('test'))
print(b)