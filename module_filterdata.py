import numpy as np
import pandas as pd


# 2022-09-22
#功能描述: 进行数据的过滤和清洗；剔除空行;标注统计标记；拆分生成主类，子类列
def filter_data(df):
    # 首先将Duration处理只保留2位小鼠
    df["Duration"] = round(df["Duration"], 2)

    """
   #1.dropna(thresh=3)剔除出NaN的空字段；必须保证有3个非空字段；也就是3个字段都必须有值
   #2.但是这种方式不是不太安全；如果抽取的是N列；有时候有些空列是要保留的；麻烦
   df = df.dropna(thresh=3)
   """

    # ToDO
    # 比较安全的做法是：
    # 1.将空值NaN用特俗符号替换
    df = df.fillna("#$%")

    # 2.将含有特殊符号的行过滤掉；
    #   同时重置索引；否则还是保持原始表中索引序列;后面添加新列数据时索引不能匹配
    df = df[df["Employee"] != "#$%"].reset_index(drop=True)

    # 获取统计年月,原始格式是12/01/2021; 按'/'拆分成列表['12','01','2021']
    # 然后只取其中的年,月，拼接成新的字符串 '2021-12'
    curDate = df.at[1, "Activity Date"].split("/")
    curDate = curDate[2] + "-" + curDate[1]

    # 定义两个List用来存放拆分Product之后的主类和子类
    lst_Class = []
    lst_SubClass = []

    # 按照索引号遍历；对Product字段按照':'进行拆分；获取主类和子类名称
    # 分别存入lst_Class，lst_SubClass两个List中
    for i in df.index:
        # 按照索引号,取得当前行的Product名称
        product = df.at[i, "Product"]
        # 获取当前Product名称中':'所在位置;用于截取主类和子类
        strIndex = product.find(":")
        # 获取当前行的统计标记
        flag = df.at[i, "Flag"]

        # 如果统计标记为1；那么就需要将Product拆分出主类和子类
        if flag == 1:
            # 有一些Product名称不包含‘:’,这是返回的位置是-1
            # 如果':'的位置不等于-1,进行拆分
            if strIndex != -1:
                strClass = product[0:strIndex]
                strSubClass = product[strIndex + 1:]
            else:
                # 如果'：'位置为-1；这种情况就不要拆分了; 主类就是子类
                strClass = product
                strSubClass = product
        else:
            # 如果统计标记为2，拆分出来之后；让子类等于主类
            # 这样统计的时候；二级分组就不起作用了
            if strIndex != -1:
                strClass = product[0:strIndex]
                strSubClass = product[0:strIndex]
            else:
                strClass = product
                strSubClass = product
        # 当前行处理获得的主类和子类名；分别追加各自的List中
        lst_Class.append(strClass)
        lst_SubClass.append(strSubClass)

    # 将主类和子类两个List的值分别添加到DF的主类和子类字段中
    df["Class"] = lst_Class
    df["SubClass"] = lst_SubClass

    # 只选取字段['Flag', 'Class', 'SubClass', 'Employee', 'Duration']
    df = df.iloc[:, [0, 1, 2, 4, 6]]

    # 返回DF数据集和统计日期
    return [df, curDate]