
from  module_filterdata import filter_data
from  module_groupby import groupby_data
from  module_loaddata import load_data




def main():

    """
    TODO

    总体思路2022-09-22：
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
    res1 = load_data(monthFile, flagSN)

    # 调用数据过滤函数；对装载的数据进行过滤清洗
    res2 = filter_data(res1)
    # 调用分组统计函数进行统计并输入结果到Excel表
    groupby_data(res2)


if __name__ == "__main__":
    main()
