import numpy as np
import pandas as pd

# 2022-09-22
#功能描述：上传Excel文件; 做初步处理生成DataFrame

def loadData(monthFile, flagSN):
    # 从Excle文件导入的数据只提取4列(Employee,Activity Date,Product/Service,Duration)
    df = pd.read_excel(monthFile, header=4).iloc[:, [1, 2, 3, 6]]

    # 列名太长不方便；改短
    df.rename(columns={"Product/Service": "Product"}, inplace=True)

    # TODo
    # 1.插入3个空列 Flag Class，Subclass（统计标识，大类，子类）;
    # 2.未赋值之前先以空值NaN代替
    df.insert(0, "Flag", value=np.nan)
    df.insert(1, "Class", value=np.nan)
    df.insert(2, "SubClass", value=np.nan)

    # 重新设置只进行主类统计的起始行
    # 总行数 - 表头4行 - 标题占1行 - 索引从0开始减一行
    flagSN = flagSN - 4 - 1 - 1

    # 根据统计标记起始行号flagSN；对每一个行的Flag进行标记
    # 小于标记行号的统一标记为1;否则标记为2
    # 这样标记1需要统计大类和小类；标记2就只统计大类
    for i in df.index:
        if i < flagSN:
            df.at[i, "Flag"] = 1
        else:
            df.at[i, "Flag"] = 2
    # 返回结果集DataFrame
    return df
