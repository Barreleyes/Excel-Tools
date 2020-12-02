import os
import argparse
import sheet_load
import current
import localization



def path_check(*path_list):
    '''
    检查目标路径是否存在，若不存在则自动创建
    '''
    for path in path_list:
        if not os.path.exists(path):
            os.makedirs(path)

def save_file(path,data):
    f = open(path, 'w', encoding='utf-8')
    f.write(data)
    f.close
# 导出配置表sheet


def export_sheets(sheet_list):
    for sheet in sheet_list:
        sheet_load.load_sheet(current.Workbook.obj.sheet_by_name(sheet))
        save_file(current.file_name('client'),current.output_str('client'))
        save_file(current.file_name('server'),current.output_str('server'))

if __name__ == "__main__":
    # 获取外部参数
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--input', default='.\\test_xlsx\\',
                    help="excel文件或文件夹路径，默认读取【工具目录】")
    parser.add_argument('-oc', '--output-client', default='.\\export\\client\\',
                    help="前端导出文件路径，默认在【工具目录\\export\\client】文件夹")
    parser.add_argument('-os', '--output-server', default='.\\export\\server\\',
                    help="后端导出文件路径，默认在【工具目录\\export\\server】文件夹")
    parser.add_argument('-fc', '--format-client', default='lua',
                    help="前端导出格式，可根据项目定制，默认为lua")
    parser.add_argument('-fs', '--format-server', default='json',
                    help="后端导出格式，可根据项目定制，默认为lua")
    parser.add_argument('-ver', '--version', default='0.0.0.1',
                    help="配置版本，会在导出文件中添加配置表版本信息")
    # 初始化外部参数
    current.Work.basic_info(parser.parse_args())
    # 检查导出路径
    path_check(current.Work.OutputPath.client,
            current.Work.OutputPath.server)
    # 导出配置表
    for name,path in current.Work.book_list.items():
        current.Workbook.basic_info(name)
        export_sheets(current.Workbook.sheet_list_normal)
        export_sheets(current.Workbook.sheet_list_lite)
