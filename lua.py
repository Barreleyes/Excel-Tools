head="--ENCODING: UTF-8--\nreturn{\n"
tail="}"
file_format='.lua'
output_data=''
def convert_data(depth,input_data):
    global output_data
    for key,val in input_data.items():
        output_data += '\t'*depth
        if isinstance(val,dict):
            if isinstance(key,int):
                output_data += '[%s] = ' % str(key)
            elif isinstance(key, str):
                output_data += '[\"%s\"] = ' % str(key)
        elif isinstance(key,str):
            output_data += '%s = ' % key

        if isinstance(val, dict):
            output_data += '{\n'
            convert_data(depth+1, val)
            output_data += '\t'*depth
            output_data += '},'
        elif isinstance(val, (int,float)):
            output_data += '%s,' % str(val)
        elif isinstance(val, str):
            output_data += '\"%s\",' % val
        elif isinstance(val,list):
            output_data += str(val).replace('[','{').replace(']','}') + ","
        output_data += '\n'
def dump(data):
    global output_data
    output_data=''
    convert_data(1,data)
    return head + output_data + tail
