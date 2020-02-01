file_format='.lua'
__head="--ENCODING: UTF-8--\nreturn{\n"
__tail="}"
__temp_obj = None
def dump(data):
    def dump_function(data):
        global __temp_obj
        __temp_obj += data
    global __temp_obj
    __temp_obj=''
    convert_data(1,data,dump_function)
    return __head + __temp_obj + __tail
def dump_file(data,output_path,file_name):
    global __temp_obj
    def dump_function(data):
        global __temp_obj
        __temp_obj.write(data)
    __temp_obj=open(output_path + '\\' + file_name + file_format,'w',encoding='utf-8')
    __temp_obj.write(__head)
    convert_data(1,data,dump_function)
    __temp_obj.write(__tail)
    __temp_obj.close()
def convert_data(depth,input_data,dump_function):
    '''
    递归处理python字典对象
    '''
    for key,val in input_data.items():
        dump_function('\t'*depth)
        if isinstance(val,dict):
            if isinstance(key,int):
                dump_function('[%s] = ' % str(key))
            elif isinstance(key, str):
                dump_function('[\"%s\"] = ' % str(key))
        elif isinstance(key,str):
            dump_function('%s = ' % key)
        elif isinstance(key,int):
            dump_function('[%s] = ' % str(key))
        if isinstance(val, dict):
            dump_function('{\n')
            convert_data(depth+1, val,dump_function)
            dump_function('\t'*depth)
            dump_function('},')
        elif isinstance(val, (int,float)):
            dump_function('%s,' % str(val))
        elif isinstance(val, str):
            dump_function('\"%s\",' % val)
        elif isinstance(val,list):
            dump_function(str(val).replace('[','{').replace(']','}') + ",")
        dump_function('\n')
