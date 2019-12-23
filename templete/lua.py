def var_number(data):
    return data
def var_string(data):
    return '\"' + str(data) + '\"'
def var_text(data):
    return '\"' + str(data) + '\"'
def array_number(data):
    output=''
    for i in data.split(ARRAY_SPLITER_D1):
        element=''
        for j in i.split(ARRAY_SPLITER_D2):
            element+=str(var_number(j)) + ','
        output+='{'+element[:-1] +'},'
    return '{' + output[:-1] + '}'
def array_string(data):
    output=''
    for i in data.split(ARRAY_SPLITER_D1):
        element=''
        for j in i.split(ARRAY_SPLITER_D2):
            element+='\"' + var_string(j) + '\",'
        output+='{'+element[:-1] +'},'
    return '{' + output[:-1] + '}'
def array_text(data):
    return array_string(data)
def dictionary(data):
    import re
    output=''
    for element in data.split(DICT_SPLITER_D1):
        part=re.search(r'\[(.+)\](.+)=(.+)',element)
        field_type=part.group(1)
        field_name='\"' +part.group(2) + '\"' if not part.group(2).isdigit() else part.group(2)
        filed_value=part.group(3)
        output+='[{id}]={data},\n'.format(id=field_name,data=str(FIELD_TYPE[field_type](filed_value))) 
    return '{' + output + '}'
FIELD_TYPE={
    "n":var_number,
    "s":var_string,
    "t":var_text,
    "an":array_number,
    "as":array_string,
    "at":array_text,
    "d":dictionary
}

ITEM='[{id}]={{{data}}},\n'
HEAD="return{\n"
TAIL="\n}"
ARRAY_SPLITER_D1='|'
ARRAY_SPLITER_D2=';'
DICT_SPLITER_D1='|'
DICT_SPLITER_D2='='

def value_format(head_info,value):
    value_name=head_info['name']
    value_type=head_info['type']
    value_default=head_info['default']
    if value=='':
        value=value_default
    formated_text=value_name +'=' + str(FIELD_TYPE[value_type](value))
    return formated_text + ';\n'
test='1;2;3|4;5;6'