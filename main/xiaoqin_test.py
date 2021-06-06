from Tool.OperationExcel import *
import os, re

rExcel = rExcel()
wExcle = wExcle()
listDataHead = ['序号', '手术日期', '患者姓名', '患者年龄', '住院号', '手术医师', '区', '床位号', '产品名称（实际使用）', '产品描述', '规格型号（实际使用）', '批号',
                '储存条件', '生产日期', '失效日期', '注册证编号', '生产厂商', '生产许可证号', '数量（实际使用）', '单价（底价）', '底价成本', '产品编码（医院计费）',
                '产品名称（医院计费）', '规格型号', '产品批号', '数量', '单价', '小计', '订单金额', '华润7%', '销售额', '跟台费', '消毒费', '餐费', '业务提成',
                '台数', '临床30%', '分销总价', '分销利润', '分销单价']


# 写入结果文件
def wData(sheet):
    with open('../TXT/{}.txt'.format(realTime()), "r") as f:  # 打开结果文件
        flag = 1  # 数据条数判定
        for i in f:
            strdata = strAddList(i)
            if "病人姓名" in strdata[0]:
                listdata = strdata[0].strip().split(' ')
                for j in listdata:
                    if '姓名' in j:
                        wExcle.wDataCell(sheet=sheet, row=flag + 1, column=listDataHead.index('患者姓名') + 1, data=j[5:])
                    if '病室' in j:
                        wExcle.wDataCell(sheet=sheet, row=flag + 1, column=listDataHead.index('区') + 1,
                                         data=re.findall(r'\d+', j)[0])
                    if '床号' in j:
                        wExcle.wDataCell(sheet=sheet, row=flag + 1, column=listDataHead.index('床位号') + 1,
                                         data=re.findall(r'\d+', j)[0])
                    if '年龄' in j:
                        wExcle.wDataCell(sheet=sheet, row=flag + 1, column=listDataHead.index('患者年龄') + 1,
                                         data=re.findall(r'\d+', j)[0])
                    if '日期' in j:
                        wExcle.wDataCell(sheet=sheet, row=flag + 1, column=listDataHead.index('手术日期') + 1,
                                         data=re.findall(r'\d+', i)[0] + '/' + re.findall(r'\d+', i)[1] +
                                              '/' + re.findall(r'\d+', j)[2])
            else:
                flag = flag + 1
                if strdata[2] != '':
                    wExcle.wDataCell(sheet=sheet, row=flag + 1, column=listDataHead.index('产品名称（医院计费）') + 1, data=strdata[2])
                if strdata[3] != '':
                    wExcle.wDataCell(sheet=sheet, row=flag + 1, column=listDataHead.index('规格型号（实际使用）') + 1, data=strdata[3])
                if strdata[4] != '':
                    wExcle.wDataCell(sheet=sheet, row=flag + 1, column=listDataHead.index('批号') + 1, data=strdata[4])
                if strdata[10] != '':
                    wExcle.wDataCell(sheet=sheet, row=flag + 1, column=listDataHead.index('生产日期') + 1, data=strdata[10])
                if strdata[5] != '':
                    wExcle.wDataCell(sheet=sheet, row=flag + 1, column=listDataHead.index('失效日期') + 1, data=strdata[5])
                if strdata[11] != '':
                    wExcle.wDataCell(sheet=sheet, row=flag + 1, column=listDataHead.index('注册证编号') + 1, data=strdata[11])
                if strdata[8] != '':
                    wExcle.wDataCell(sheet=sheet, row=flag + 1, column=listDataHead.index('数量（实际使用）') + 1, data=strdata[8])
                if strdata[7] != '':
                    wExcle.wDataCell(sheet=sheet, row=flag + 1, column=listDataHead.index('单价') + 1, data=strdata[7])
                if strdata[9] != '':
                    wExcle.wDataCell(sheet=sheet, row=flag + 1, column=listDataHead.index('小计') + 1, data=strdata[9])


# 判断是否有当月sheet , 没有则创建再写数据
def judgeSheet(path):
    work = wExcle.openExcel(path)
    sheetNames = wExcle.getSheetName(work)
    if realTime() not in sheetNames:
        sheet = wExcle.creatSheet(work)
        wExcle.newRemplate(sheet)
        wData(sheet)
        work.save(path)
    else:
        sheet = wExcle.checkOutSheet(work, realTime())
        wData(sheet)
        work.save(path)


# 读取源数据
list_filePath = dataPath(getPath() + '\\DATA\\{}\\'.format(realTime()),
                         getFile("../DATA/{}/".format(realTime())))  # 将数据文件路径放入一个列表中
for i in list_filePath:  # 批量读取数据文件并保存到txt文档
    wb = rExcel.openExcle(i)  # excel实例对象
    sheetNames = rExcel.getSheetNames(wb)  # 获取所有sheet名
    sheet = rExcel.checkOutSheet(wb, sheetNames[0])  # 切换到第一个sheet表
    list_data = rExcel.getRowData(sheet)  # 获取该sheet表中所有行的数据
    if os.path.exists("../TXT"):  # 判断是否有TXT文件夹
        wTxt('../TXT/{}.txt'.format(realTime()), list_data)
    else:
        os.makedirs("../TXT")
        wTxt('../TXT/{}.txt'.format(realTime()), list_data)
# 写到另一个exlce表中
if os.path.exists('../result/植入单-中南大学湘雅三医院.xlsx'):  # 判断输入表是否存在
    judgeSheet('../result/植入单-中南大学湘雅三医院.xlsx')
else:
    wExcle.creatExcel('../result/植入单-中南大学湘雅三医院.xlsx')
    judgeSheet('../result/植入单-中南大学湘雅三医院.xlsx')
