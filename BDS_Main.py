import numpy as np
import pandas as pd


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


def filterData(df):
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
                strSubClass = product[strIndex + 1 :]
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


def groupBy(res_Filter):
    df = res_Filter[0]  # 获取数据集
    curdate = res_Filter[1]  # 获取统计日期

    # 按照按主类，子类，员工分组统计字段Duration(即每个员工的合计)
    df_Group_Class_SubClass_Employee = df.groupby(["Class", "SubClass", "Employee"])[
        "Duration"
    ].sum()

    # 按照主类，子类分组统计字段Duration(即每个子类的合计)
    df_Group_Class_SubClass = df.groupby(["Class", "SubClass"])["Duration"].sum()

    # 生成新的DataFrame
    lst_Class = []  # 主类
    lst_SubClass = []  # 子类
    lst_Employee = []  # 员工
    lst_Duration = []  # 合计工时
    lst_Percentage = []  # 所在百分比

    # 按照员工分组汇总结果进行遍历;; 这是联合索引('Class','SubClass','Employee')
    for i in df_Group_Class_SubClass_Employee.index:
        # 取出当前员工的工时数
        sumEmployee = df_Group_Class_SubClass_Employee.loc[i]
        # 按照主类和子类;获取子类的工时小计
        sumClass = df_Group_Class_SubClass.loc[(i[0], i[1])]
        # 计算出当前员工工时所占子类工时的百分比
        percentage = str(round((sumEmployee / sumClass) * 100, 2)) + "%"

        # 将结果分别追加到5个List中
        lst_Percentage.append(percentage)
        lst_Class.append(i[0])
        lst_SubClass.append(i[1])
        lst_Employee.append(i[2])
        lst_Duration.append(sumEmployee)

    # 分别生成工时和百分比两个统计结果集; 数据组织必须用字典(Dict)格式
    # 注：(可以将工时和百分比生成为一个结果集；但是无法对两个值进行旋转)
    df_Duration = pd.DataFrame(
        {
            "Class": lst_Class,
            "SubClass": lst_SubClass,
            "Employee": lst_Employee,
            "Duration": lst_Duration,
        }
    )

    df_Percentage = pd.DataFrame(
        {
            "Class": lst_Class,
            "SubClass": lst_SubClass,
            "Employee": lst_Employee,
            "Percentage": lst_Percentage,
        }
    )

    # 按照按主类，子类两级索引；将员工-Duration行转列
    df_index_Duration = df_Duration.set_index(["Class", "SubClass", "Employee"])[
        "Duration"
    ]
    df_index_Duration = df_index_Duration.unstack()

    # 按照按主类，子类两级索引；将员工-Percentage行转列
    df_index_Percentage = df_Percentage.set_index(["Class", "SubClass", "Employee"])[
        "Percentage"
    ]
    df_index_Percentage = df_index_Percentage.unstack()
    # 保存为Excel文件；尝试将两个数据集作为两个Sheet保存到一个Excel文件中；未成功
    # 1.结果不含百分比的文件
    fileName1 = "d:/QQDownload/BDS-Duration-" + curdate + ".xlsx"
    df_index_Duration.to_excel(fileName1)

    # 2.结果含有百分比的文件
    fileName2 = "d:/QQDownload/BDS-Percentage-" + curdate + ".xlsx"
    df_index_Percentage.to_excel(fileName2)


def main():

    """
    TODO
    总体思路：
    1.两种统计方式；【大类+小类+人员】 【大类+人员】
    2.用一个Flag列来标记统计类别
      将【大类+小类+人员】统计的行标记为1
      将【大类+人员】统计的行标记为2
    3.将Product/Service列用':'拆分成大类和小类
      如果标记为1,就分别拆分出大类和小类
      如果标记为2,只拆分出大类; 让小类也等于大类;
      这样小类分组统计事实上无意义;也就实现了只按大类统计
    4.按照【大类+小类+人员】分组统计出每个人员工时
      按照【大类+小类】分组统计出每个小类的工时总计
      计算出每个人员工时所占小类百分比
    5.将结果按照统计日期保存为Excel文件
    6.分为三个函数实现：
      def loadData(monthFile,flagSN) 上传Excel文件; 做初步处理生成DataFrame
      def filterData(df) 进行数据的过滤和清洗；剔除空行;标注统计标记；拆分生成主类，子类列
      def groupBy(res_Filter) 分组统计；并将结果保存为Excel文件
    """

    # 指定要处理的Excel文件
    monthFile = "d:/QQDownload/newBDS.xlsx"
    # 指定从哪一行开始只统计大类;不统计小类
    flagSN = 338

    # 调用数据装载函数；返回一个DataFrame
    # 返回的列：['Flag', 'Class', 'SubClass', 'Activity Date', 'Employee', 'Product', 'Duration']
    res1 = loadData(monthFile, flagSN)

    # 调用数据过滤函数；对装载的数据进行过滤清洗
    res2 = filterData(res1)
    # 调用分组统计函数进行统计并输入结果到Excel表
    groupBy(res2)


if __name__ == "__main__":
    main()
