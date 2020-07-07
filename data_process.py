import math
array_spliter_d1='|'
array_spliter_d2=';'
dict_spliter_d1='|'
dict_spliter_d2='='
def array_process(data,var_type):
    new_data=[]
    if array_spliter_d2 in data:
        for i in data.split(array_spliter_d1):
            if i !='':
                temp=[]
                for j in i.split(array_spliter_d2):
                    temp.append(var_type(j))
                new_data.append(temp)
    else:
        for i in data.split(dict_spliter_d1):
            if i !='':
                new_data.append(var_type(i))
    return new_data
def var_int(data,count=1):
    return int(data)
def var_float(data,count=1):
    return int(float(data)*1000000)
def var_number(data,count=1):
    return int(data) if math.modf(float(data))[0] ==0 else float(data)
def var_string(data,count=1):
    return str(data)
def var_text(data,count=1):
    return str(data)
def array_int(data,count=1):
    return array_process(data,var_int)
def array_float(data,count=1):
    return array_process(data,var_float)
def array_number(data,count=1):
    return array_process(data,var_number)
def array_string(data,count=1):
    return array_process(data,str)
def array_text(data,count=1):
    return array_process(data,var_text)    
def key_int(data,count=1):
    return int(data)
def key_string(data,count=1):
    return str(data)
def dictionary(data,count=1):
    import re
    output={}
    for element in data.split(dict_spliter_d1):
        part=re.search(r'\[(.+)\](.+)'+ dict_spliter_d2 +'(.+)',element)
        field_type=part.group(1)
        field_name=int(part.group(2)) if part.group(2).isdigit() else part.group(2)
        filed_value=part.group(3)
        output[field_name]=process_type[field_type](filed_value)
    return output 
process_type={
    "i":var_int,
    "f":var_float,
    "n":var_number,
    "s":var_string,
    "t":var_text,
    "ai":array_int,
    "af":array_float,
    "an":array_number,
    "as":array_string,
    "at":array_text,
    "d":dictionary,
    "ikey":key_int,
    "skey":key_string,
}