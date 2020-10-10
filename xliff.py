
# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

import keyword
import xlrd
import xml.etree.ElementTree as ET
from xml.dom import minidom

# 解析xliff的nameSpace Xliff版本变更请更改这里
ns = dict(xliffNameSpace='urn:oasis:names:tc:xliff:document:1.2')
# excel地址
excelPath = r'/Users/cc/Downloads/One所有文案 (2).xlsx'

# xliff地址
xliffPath = r'/Users/cc/Desktop/UDictionary/id.xcloc/Localized Contents/id.xliff'

def readExcel():
    data = xlrd.open_workbook(excelPath)
    # print(data.sheet_names())
    # 拿到索引 默认是在第一个
    sheet1 = data.sheet_by_index(0)
    # print('索引名称：' + str(sheet1.name) + ' 索引的行数' + str(sheet1.nrows) + ' 索引的列数' + str(sheet1.ncols))
    # row_value = sheet1.row_values(2)
    # print(row_value)
    # col_value = sheet1.col_values(2)
    # print(col_value)
    # print(sheet1.nrows)
    # 在这里修改语言
    sourceName = "English"
    targetName = "Indonesian"

    sourceIndex = 0
    targetIndex = 0
    # 确定Index
    row_0 = sheet1.row_values(0)
    for k in range(len(row_0)):
        if row_0[k] == sourceName:
            sourceIndex = k
            print(sourceName + "所在索引" + str(k))
        elif row_0[k] == targetName:
            targetIndex = k
            print(targetName + "所在索引" + str(k))
    # 简历字典
    dict = {}
    for i in range(sheet1.nrows):
        row_value = sheet1.cell_value(i, targetIndex)
        row_key = sheet1.cell_value(i, sourceIndex)
        dict[row_key] = row_value
    # print(dict)
    #读取Xliff
    readXliff(dict)



def readXliff(dict):
    # 重写nameSpace
    ET.register_namespace('','urn:oasis:names:tc:xliff:document:1.2')
    #拿到根节点
    tree = ET.parse(xliffPath)
    root = tree.getroot()
    # print(root.tag)
    for file in root.findall('xliffNameSpace:file', ns):
        body = file.find('xliffNameSpace:body', ns)
        for unit in body.findall('xliffNameSpace:trans-unit', ns):
            source = unit.find('xliffNameSpace:source', ns)
            # print(source.text)
            target = unit.find('xliffNameSpace:target', ns)
            #target不存在就创建 在excel中填写疑问 否则写入原文
            if target is None:
                #创建target
                node = ET.SubElement(unit, "target")
                if source.text in dict:
                    node.text = dict[source.text]
                else:
                    node.text = source.text
                node.tail = '\n\t'
                print('create success')
            elif source.text in dict:
                target.text = dict[source.text]
    saveXML(root, xliffPath)
    print('write success')


def subElement(root, tag, text):
    ele = ET.SubElement(root, tag)
    ele.text = text

#不再换行
def saveXML(root, filename, indent="\t", newl="", encoding="utf-8"):
    rawText = ET.tostring(root)
    dom = minidom.parseString(rawText)
    with open(filename, 'w') as f:
        dom.writexml(f, "", indent, newl, encoding)

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    readExcel()
    # readXliff()
