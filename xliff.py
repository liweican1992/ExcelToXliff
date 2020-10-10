# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

import keyword
import xlrd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import os

# 解析xliff的nameSpace Xliff版本变更请更改这里
ns = dict(xliffNameSpace='urn:oasis:names:tc:xliff:document:1.2')

# excel地址
excelPath = r'/Users/cc/Downloads/One所有文案.xlsx'

# xliff地址
xliffPath = r'/Users/cc/Desktop/UDictionary/id.xcloc/Localized Contents/id.xliff'


def readExcel():
    if not os.path.exists(excelPath):
        print('excel路径不存在')
        return
    if not os.path.exists(xliffPath):
        print('xliff路径不存在')
        return
    data = xlrd.open_workbook(excelPath)
    # 拿到索引 默认是在sheet0
    sheet1 = data.sheet_by_index(0)

    # 在这里修改语言
    sourceName = "English"
    targetName = "Indonesian"
    sourceIndex = -1
    targetIndex = -1
    # 确定Index
    row_0 = sheet1.row_values(0)
    for k in range(len(row_0)):
        if row_0[k] == sourceName:
            sourceIndex = k
            print(sourceName + " 所在索引 " + str(k))
        elif row_0[k] == targetName:
            targetIndex = k
            print(targetName + " 所在索引 " + str(k))

    if sourceIndex == -1:
        print('未找到English Index')
        return
    if targetIndex == -1:
        print('未找到 ' + targetName + ' Index')
        return
    # 遍历字典
    dict = {}
    for i in range(sheet1.nrows):
        row_value = sheet1.cell_value(i, targetIndex)
        row_key = sheet1.cell_value(i, sourceIndex)
        dict[row_key] = row_value
    print('excel数据读取完毕，共 ', len(dict), ' 条数据')
    # 读取Xliff
    writeXliff(dict)


def writeXliff(dict):
    print('读取xliff文件')
    # 重写nameSpace
    ET.register_namespace('', 'urn:oasis:names:tc:xliff:document:1.2')
    # 拿到根节点
    tree = ET.parse(xliffPath)
    root = tree.getroot()
    # print(root.tag)
    count = 0
    for file in root.findall('xliffNameSpace:file', ns):
        print('读取子文件：' + file.attrib['original'])
        body = file.find('xliffNameSpace:body', ns)
        for unit in body.findall('xliffNameSpace:trans-unit', ns):
            source = unit.find('xliffNameSpace:source', ns)
            # print(source.text)
            target = unit.find('xliffNameSpace:target', ns)
            # target不存在就创建 在excel中填写译文 否则写入原文
            if target is None:
                # 创建target
                node = ET.SubElement(unit, "target")
                if source.text in dict:
                    node.text = dict[source.text]
                    count += 1
                else:
                    node.text = source.text
                    print('填补空白译文')
                node.tail = '\n\t'
                print('create success')
            elif source.text in dict:
                target.text = dict[source.text]
                count += 1

    print('写入完成，共写入 ' + str(count) + ' 条数据')
    saveXML(root, xliffPath)
    print('finish')


def subElement(root, tag, text):
    ele = ET.SubElement(root, tag)
    ele.text = text


# 格式优化 不再换行
def saveXML(root, filename, indent="\t", newl="", encoding="utf-8"):
    rawText = ET.tostring(root)
    dom = minidom.parseString(rawText)
    with open(filename, 'w') as f:
        dom.writexml(f, "", indent, newl, encoding)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    readExcel()
    # readXliff()
