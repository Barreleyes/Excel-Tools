class TableType:
    class Normal:
        DATA_ROW = 2
        HEAD_ATTRIBUTE = ['name_cn', 'name', 'stage', 'default']

    class Config:
        DATA_ROW = 2
        HEAD_ATTRIBUTE = ['name_cn', 'name', 'stage', 'default']


class FieldInfo:
    name = ''
    ftype = ''
    count = 1
    stage = ''
    default = ''
    column = 0


VALID_INPUT_FORMAT = ['xlsx', 'xlsm', 'xls']


class Current:
    class Work:
        input = ''

        class OutputPath:
            client = ''
            server = ''

        class OutputFormat:
            client = ''
            server = ''

        class DumpMethod:
            client=None
            server = None
        book_list = []
        @staticmethod
        def set(args):
            Current.Work.input = args.input
            Current.Work.OutputPath.client = args.output_client
            Current.Work.OutputPath.server = args.output_server
            Current.Work.OutputFormat.client = args.format_client
            Current.Work.OutputFormat.server = args.format_server

    class Workbook:
        name = ''
        sheets = []
        path = []
        sheet_list_normal = {}
        sheet_list_config = {}

    class Sheet:
        name = ''
        sheet_type = ''
        row_max = 0
        col_max = 0
        mainkey_col = None
        subkey_col = None
        head_info = {}

    class Document:
        mainkey = ''
        subkey = ''

    class Cell:
        address = [0, 0]
        @staticmethod
        @property
        def row():
            return Current.Cell.address[0]
        @staticmethod
        @property
        def column():
            return Current.Cell.address[1]
        @staticmethod
        def set(row,col):
            Current.Cell.address=[row+1,col+1]

    class Data:
        value = ''
        vtype = ''
        index = ''
    Field = FieldInfo()
    @staticmethod
    def exception_str():
        return "【表名：%s；sheet名：%s；字段名：%s；主键：%s；子键：%s；行：%s；列：%s；】" % \
            (Current.Workbook.name, Current.Sheet.name, Current.Field.name,
             Current.Document.mainkey, Current.Document.subkey, Current.Cell.row, Current.Cell.column)

    @staticmethod
    def key_str(index=None):
        if index == None:
            return "%s_%s_%s_%s_%s" % \
                (Current.Workbook.name, Current.Sheet.name,
                    Current.Field.name, Current.Document.mainkey, Current.Document.subkey)
        else:
            return "%s_%s_%s_%s_%s_%s" % \
                (Current.Workbook.name, Current.Sheet.name,
                    Current.Field.name, Current.Document.mainkey, Current.Document.subkey, ''.join(map(str, index)))

    @staticmethod
    def file_name(stage: str):
        if stage == 'client':
            path, oformat = Current.Work.OutputPath.client, Current.Work.OutputFormat.client
        elif stage == 'server':
            path, oformat = Current.Work.OutputPath.server, Current.Work.OutputFormat.server
        else:
            raise Exception("平台类型参数错误")
        return '%s\\%s_%s.%s' % (path, Current.Workbook.name, Current.Sheet.name, oformat)

    @staticmethod
    def clear():
        Current.Workbook
