import importlib
import xlrd
import sys,os
import data_checker
import re
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
            vertical.append(sheet.name)
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
    key_info=[]
    mainkey_col=0
    subkey_col=0
    key_count=0
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
            head_info[i]['default']=default_value_cell
            head_info[i]['platform']=field_range_cell
            if head_info[i]['type'][0]=='a':
                head_info[i]['count']=int(head_info[i]['type'][-1])
                head_info[i]['type']=part.group(1)[:-1]
            else:
                head_info[i]['count']=0
                pass
            if 'key' in head_info[i]['type']:
                key_count+=1
                if key_count==1:
                    mainkey_col=i
                elif key_count==2:
                    subkey_col=i
                    head_info[i]['type']+='_sub'
    
    if key_count==0:
        raise Exception("单张配置表应至少有一个主键")
    elif key_count==2:
        start=4
        for row in range(5,sheet.nrows):
            if sheet.cell(row,0).value !='' or row == sheet.nrows-1:
                if row==sheet.nrows-1:
                    key_info.append([start,row])
                else:
                    key_info.append([start,row-1])
                    start=row
    elif key_count>2:
        raise Exception("单张配置表最多有两个键")
    return head_info,mainkey_col,subkey_col,key_info  
def gen_export_file_horizon(sheet):
    def item_process(id,row,col_start,col_end,head_info):
        '''
        :param id: 配置id
        :param row: 配置所在行
        :param col_start: 导出起始列
        :param col_end: 导出结束列
        :param head_info: 表头信息
        导出配置数据
        '''
        # 判断ID是整数还是字符串
        item_id= '\"' + str(id) +'\"' if isinstance(id,str) else int(id)
        item_data=''
        for column in range(col_start,col_end):
            if head_info[column]:
                item_data+=EXPORT_LANG.dump(head_info[column],sheet.cell(row,column).value)
        # 预留子key的替换项
        return EXPORT_LANG.ITEM.format(id=item_id,data=item_data)
    #获取表头信息
    head_info,mainkey_col,subkey_col,key_info=get_head_info(sheet)
    # 生成导出文件路径及文件名
    lua_name=export_path + '\\' + WORKBOOK_NAME + '_' + sheet.name + EXPORT_LANG.FILE_TYPE
    # 生成导出文件
    export_file=open(lua_name,'w',encoding='utf-8')
    # 写入文件头
    export_file.write(EXPORT_LANG.HEAD)
    
    #导出单键表
    if subkey_col==0:
        for row in range(4,sheet.nrows):
            id=sheet.cell(row,mainkey_col).value
            if id!='':
                export_file.write(item_process(id,row,mainkey_col+1,sheet.ncols,head_info))
    # 导出双键表
    else:
        for item in key_info:
            mainkey_item=item_process(sheet.cell(item[0],mainkey_col).value,item[0],mainkey_col+1,subkey_col+1,head_info)
            subkey_item=''
            for item_sub in range(item[0],item[1]+1):
                id=sheet.cell(item_sub,subkey_col).value
                if id!='':
                    subkey_item+=item_process(id,item_sub,subkey_col+1,sheet.ncols,head_info)
            mainkey_item=mainkey_item.replace('value_text',subkey_item)
            export_file.write(mainkey_item)
    # 写入文件尾
    export_file.write(EXPORT_LANG.TAIL)
    export_file.close()
def gen_export_file_vertical(sheet):
    lua_name=export_path + '\\' + WORKBOOK_NAME + '_' + sheet.name.replace('@','') + EXPORT_LANG.FILE_TYPE
    export_file=open(lua_name,'w',encoding='utf-8')
    export_file.write(EXPORT_LANG.HEAD)
    for row in range(0,sheet.nrows):
        part=re.search(r'\[(.+)\](.+)',sheet.cell(row,1).value)
        key_info={'count':0,'default':0,'name':part.group(2),'platform':'all','type':part.group(1)}
        export_file.write(EXPORT_LANG.dump(key_info,sheet.cell(row,2).value))
    export_file.write(EXPORT_LANG.TAIL)
    export_file.close()
def path_check(path):
    if not os.path.exists(path):
        os.makedirs(path)
# 读取文件
sys.argv=['0','lua','.\\test_doublekey.xlsx','.\\export']
sys.path.append('.\\templete')
export_format=sys.argv[1]
excel_path=sys.argv[2]
export_path=sys.argv[3]
path_check(export_path)
workbook=xlrd.open_workbook(excel_path)
WORKBOOK_NAME=os.path.split(excel_path)[1].split('.')[0]
# 导入输出文件格式模板
EXPORT_LANG=importlib.import_module(export_format)
# 导出水平表和垂直表
book_vertical,book_horizon=get_sheet_list(workbook)
[gen_export_file_horizon(workbook.sheet_by_name(book)) for book in book_horizon]
[gen_export_file_vertical(workbook.sheet_by_name(book)) for book in book_vertical]