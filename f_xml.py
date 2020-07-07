#coding:utf-8
#字典转化为xml
from xml.etree.ElementTree import Element, SubElement, tostring
from xml.dom.minidom import parseString
#定义一个对象，在python中叫字典dict
myobjects = {
    'key1': {
        'key11': 'value11',
        'key12': 12,
        'key13': 13,
    },
    'key2': {
        'key21': 'value21',
        'key22': ['value22_1', 'value22_2', 'value22_3'],
        'key23': 23,
    },
    'key3': {
        'key31': 'value31',
        'key32': {1:1,2:2},
    },
}

objects = Element('objects')  #根节点
for isbn, info in myobjects.items():  #迭代每一个属性，属性和值（这里的每个值都是另一个对象）
    object = SubElement(objects, isbn)  #创建子节点
    for key, val in info.items():
        SubElement(object, key).text = ', '.join(str(val).split(':'))

xml = tostring(objects)
print('原始xml样式输出')
print(xml)

print('\n美观xml样式输出')
dom = parseString(xml)
print(dom.toprettyxml('    '))

print('遍历所有节点作为一个大的平坦结构')
for elmt in objects.getiterator():
    print(elmt.tag, '-', elmt.text)

print('\n在xml中进行简单搜索')
for book in objects.findall('.//key11'):
    print(book.text)