import os
import argparse
import data_loader
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
    # 初始工作区
    data_loader.Work.update(parser.parse_args())
    # 检查导出路径
    path_check(data_loader.Work.OutputPath.client,
               data_loader.Work.OutputPath.server)
    # 导出配置表
    for name,path in data_loader.Work.book_list.items():
        # 初始化excel文件信息
        data_loader.Workbook.update(name)
        # 导出sheet
        for sheet in data_loader.Workbook.sheet_list:
            # 初始化sheet信息
            data_loader.Sheet.update(data_loader.Workbook.obj.sheet_by_name(sheet))
            # 导出数据
            for i in range(data_loader.Sheet.data_range[0][0], data_loader.Sheet.data_range[1][0]):
                for j in range(data_loader.Sheet.data_range[0][1], data_loader.Sheet.data_range[1][1]):
                    data_loader.Cell.update(i, j)
            # 序列化sheet数据
            save_file(data_loader.file_name('client'), data_loader.output_str('client'))
            save_file(data_loader.file_name('server'), data_loader.output_str('server'))
