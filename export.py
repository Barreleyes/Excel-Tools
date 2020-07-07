import sys
import os
import argparse
import xlrd
import sheet_load
import common
import multilingual
# 创建外部参数解析器
parser = argparse.ArgumentParser()
parser.add_argument('-i', '--input', default='.\\test_doublekey.xlsx',
                    help="excel文件或文件夹路径，默认读取【工具目录】")
parser.add_argument('-oc', '--output-client', default='.\\export\\client\\',
                    help="前端导出文件路径，默认在【工具目录\\export\\client】文件夹")
parser.add_argument('-os', '--output-server', default='.\\export\\server\\',
                    help="后端导出文件路径，默认在【工具目录\\export\\server】文件夹")
parser.add_argument('-fc', '--format-client', default='lua',
                    help="前端导出格式，可根据项目定制，默认为lua")
parser.add_argument('-fs', '--format-server', default='json',
                    help="后端导出格式，可根据项目定制，默认为lua")
common.Current.Work.set(parser.parse_args())


def path_check(*path_list):
    '''
    检查目标路径是否存在，若不存在则自动创建
    '''
    for path in path_list:
        if not os.path.exists(path):
            os.makedirs(path)


def set_export_format():
    def set_format(format_type):
        if format_type == 'lua':
            import lua
            return lua.dumps
        elif format_type == 'yaml':
            import yaml
            return yaml.dump
        elif format_type == 'json':
            import json
            return json.dumps
    common.Current.Work.DumpMethod.client = set_format(
        common.Current.Work.OutputFormat.client)
    common.Current.Work.DumpMethod.server = set_format(
        common.Current.Work.OutputFormat.server)


set_export_format()
# 生成导出文件列表


def get_sheet_list(workbook: xlrd.Book):
    '''
    :param workbook:xlrd的workbook对象

    '''

    def dict_append(dst_dict, name):
        if (n:=name.split('_')[0].replace('@', '')) in dst_dict:
            dst_dict[n].append(name)
        else:
            dst_dict[n] = []
            dst_dict[n].append(name)
    for sheet in workbook.sheets():
        sheet_name = sheet.name
        if sheet_name[0] == '@':
            dict_append(common.Current.Workbook.sheet_list_config, sheet_name)
        elif sheet.name[0] == '_':
            pass
        else:
            dict_append(common.Current.Workbook.sheet_list_normal, sheet_name)


def get_file_list(path: str):
    file_list = []

    def update_filelist(path: str):
        file_list.append(path)
        common.Current.Work.book_list.append(os.path.split(path)[-1])
    if os.path.isfile(path) and path.split('.')[-1] in common.VALID_INPUT_FORMAT:
        update_filelist(path)
    elif os.path.isdir(path):
        for maindir, subdir, file_name_list in os.walk(path):
            for filename in file_name_list:
                apath = os.path.join(maindir, filename)
                if apath.split('.')[-1] in common.VALID_INPUT_FORMAT:
                    update_filelist(apath)
    return file_list


def save_file(data, path):
    f = open(path, 'w', encoding='utf-8')
    f.write(data)
    f.close
# 导出配置表sheet


def export_sheets(sheet_group, sheet_type):
    for k, v in sheet_group.items():
        key_list = []
        data_client = {}
        data_server = {}
        for sheet in v:
            common.Current.Sheet.name = k
            sheet_data = sheet_load.SheetData(
                workbook_obj.sheet_by_name(sheet), sheet_type)
            sheet_data.mainkey_list = key_list
            data_client.update(sheet_data.data_client)
            data_server.update(sheet_data.data_server)
            key_list += sheet_data.mainkey_list
        data_client = common.Current.Work.DumpMethod.client(data_client)
        data_server = common.Current.Work.DumpMethod.server(data_server)
        save_file(data_client, common.Current.file_name('client'))
        save_file(data_server, common.Current.file_name('server'))


# 导出指定目录的配置表
path_check(common.Current.Work.OutputPath.client,
           common.Current.Work.OutputPath.server)
for workbook in get_file_list(common.Current.Work.input):
    workbook_name = os.path.split(workbook)[1]
    common.Current.Workbook.name = workbook_name.split('.')[
        0]
    workbook_obj = xlrd.open_workbook(workbook)
    get_sheet_list(workbook_obj)
    export_sheets(common.Current.Workbook.sheet_list_normal, "HORIZON")
    export_sheets(common.Current.Workbook.sheet_list_config, "VERTICAL")
