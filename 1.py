import re
str1 = '病人姓名：张志龙    病室：27病区骨科   床号:11床     性别：男   年龄：57岁      住院号：730726    NO：352'
list2 = str1.strip().split(' ')
for i in list2:
    if "姓名" in i :
        print(i[5:])
    if "床号" in i:
        print(re.findall(r'\d+', i))
