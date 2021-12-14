# -*- coding: utf-8 -*-

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
# liweican@rd.netease.com

import keyword
import xlrd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import os

# 解析xliff的nameSpace Xliff版本变更请更改这里
ns = dict(xliffNameSpace='urn:oasis:names:tc:xliff:document:1.2')

# excel地址
excelPath = os.path.join(os.path.expanduser("~"), 'Desktop/localizable.xlsx')
# xliff根目录
xliffRootPath = os.path.join(os.path.expanduser("~"), 'Desktop/exportLoaclization')
# 语言list
languages = {"Spanish" : "es", "Portuguese" : "pt", "Indonesian" : "id", "Arabic" : "ar", "Japanese" : "ja", "Vietnam" : "vi-VN", "German" : "de", "French" : "fr"}

# 目标语言 需要和excel中一致
targetName = ""
sourceName = "English"

# sheetName 从哪个表读取，找不到会默认读取第一个表
sheetName = "ios"

def GetDesktopPath():
    return os.path.join(os.path.expanduser("~"), 'Desktop')

def readExcel():
   
    if not os.path.exists(excelPath):
        print('excel路径不存在')
        return

    # if not os.path.exists(xliffPath):
    #     print('xliff路径不存在')
    #     return

    data = xlrd.open_workbook(excelPath)
    # 拿到索引 默认是在sheet0
    sheet1 = data.sheet_by_index(0)
    if sheetName in data.sheet_names():
        sheet1 = data.sheet_by_name(sheetName)
    else:
        print("sheetName不存在 默认读取第一个")


    # 找English索引
    sourceIndex = -1
    # 确定Index
    row_0 = sheet1.row_values(0)
    for k in range(len(row_0)):
        if row_0[k] == sourceName:
            sourceIndex = k
            print(sourceName + " 所在索引 " + str(k))

    if sourceIndex == -1:
        print('未找到English Index')
        return


    #找target索引
    for subLan in languages:
        print(subLan)
        targetName = subLan
        targetIndex = -1
        for k in range(len(row_0)):
            if row_0[k] == targetName:
                targetIndex = k
                print(targetName + " 所在索引 " + str(k))
                # 遍历字典
                dict = {}
                for i in range(sheet1.nrows):
                    row_value = sheet1.cell_value(i, targetIndex)
                    row_key = sheet1.cell_value(i, sourceIndex)
                    if row_key == 'English' or row_key == "":
                        continue
                    dict[row_key] = row_value
                print('excel数据读取完毕，共 ', len(dict), ' 条数据')
                print(dict)
                # 读取Xliff
                writeXliff(dict,targetName)





def writeXliff(dict, targetName):
    print('========================')
    print('读取 '+targetName+' xliff文件')
    lan = languages[targetName]
    xliffPath = os.path.join(xliffRootPath, lan+'.xcloc/Localized Contents/'+lan+'.xliff')
    print(xliffPath)

    if not os.path.exists(xliffPath):
        print('xliffPath路径不存在')
        return
    print('开始解析xliff')
    # 重写nameSpace
    ET.register_namespace('', 'urn:oasis:names:tc:xliff:document:1.2')
    # 拿到根节点
    tree = ET.parse(xliffPath)
    root = tree.getroot()
    # print(root.tag)
    count = 0
    for file in root.findall('xliffNameSpace:file', ns):
        #print('读取子文件：' + file.attrib['original'])
        body = file.find('xliffNameSpace:body', ns)
        for unit in body.findall('xliffNameSpace:trans-unit', ns):
            source = unit.find('xliffNameSpace:source', ns)
            #print(source.text)
            target = unit.find('xliffNameSpace:target', ns)
            # target不存在就创建 在excel中填写译文 否则写入原文
            if target is None:
                # 创建target
                print(source.text + "target为空")
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
            elif target.text is None:
                print('*************************')
                print("target为空字符串，填充source")
                target.text = source.text
                count += 1



    print('========================')
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
