# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

import keyword
import xlrd


# def print_hi(name):
#     # Use a breakpoint in the code line below to debug your script.
#     dict = {"name":"cc", "age":22}
#     print(dict)
#     del dict["name"]
#     print(dict)

def readExcel():
    data = xlrd.open_workbook(r"/Users/cc/Desktop/One所有文案.xlsx")
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

    print(dict)


def readXML():
    print("readXML")


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # print_hi('666')
    readExcel()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
