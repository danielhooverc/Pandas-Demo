import numpy as np
import pandas as pd


#功能描述：分组统计；并将结果保存为Excel文件

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