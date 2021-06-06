from Tool.OperationExcel import *
import os

# 读取源数据
rExcel = rExcel()
list_filePath = dataPath(getPath() + '\\DATA\\{}\\'.format(realTime()), getFile("../DATA/{}/".format(realTime())))  # 将数据文件路径放入一个列表中
for i in list_filePath:  # 批量读取数据文件并保存到txt文档
    wb = rExcel.openExcle(i)  # excel实例对象
    sheetNames = rExcel.getSheetNames(wb)  # 获取所有sheet名
    sheet = rExcel.checkOutSheet(wb, sheetNames[0])  # 切换到第一个sheet表
    list_data = rExcel.getRowData(sheet)  # 获取该sheet表中所有行的数据
    print(list_data)
    print(type(list_data))
    if os.path.exists("../TXT"):  # 判断是否有TXT文件夹
        wTxt('../TXT/{}.txt'.format(realTime()), list_data)
    else:
        os.makedirs("../TXT")
        wTxt('../TXT/{}.txt'.format(realTime()), list_data)
# 写到另一个exlce表中
wExcle = wExcle()
if os.path.exists('../result/植入单-中南大学湘雅三医院.xlsx'):  # 判断输入表是否存在
    pass
