import os, time


# 获取项目路径
def getPath():
    path = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
    return path


# 获取当前时间 , 年月日
def realTime():
    return time.strftime("%Y-%m")


# 返回指定目录下所有的文件
def getFile(path):
    list_file = os.listdir(path=path)
    return list_file


# 拼接数据路径
def dataPath(path, list_file):
    list_path = []
    for i in list_file:
        data = os.path.join(path, i)
        # data = data.replace('\\','/')
        list_path.append(data)
    return list_path


# 写入数据到txt里面
def wTxt(path, data):
    with open(path, "a") as f:
        if isinstance(data , list):
            for i in data:
                f.write(i)
        else:
            f.write(data)


if __name__ == '__main__':
    print(getPath())
    print(realTime())
    print(getFile("../DATA/"))
    print(dataPath(getPath() + '\\DATE\\', getFile("../DATA/{}/".format(realTime()))))
    print("../DATA/{}/".format(realTime()))
