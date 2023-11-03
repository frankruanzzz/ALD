import pandas as pd
import numpy as np

date = input("输入日期_")


path_112 = f"F:\RYH\物料相关\每日数据源\{date}\stock.quant_112_PM.xlsx"
Output_path_112_计划工作台用 = f"F:\RYH\物料相关\每日数据源\{date}\stock.quant_112_OK_{date}_晚核对工单用.xlsx"

df_112 = pd.read_excel(path_112)
df_112.loc[:,"不带C项目"] = df_112["产品/内部参考"].str[:7]
df_112.loc[:,"项目号末尾"] = df_112["产品/内部参考"].str[-1]
df_112.loc[:,"数量"]= df_112.loc[:,"数量"].astype("float")
def Condition_112(df_112):
    return  (df_112.loc[:,"位置"]!="YP-DJ")&\
            (df_112.loc[:,"位置"]!="YP-DJ")&\
            (df_112.loc[:,"位置"]!="YP-HG")&\
            (df_112.loc[:,"位置"]!="YP-HG2-8")&\
            (df_112.loc[:,"位置"]!="BLYF")&\
            (df_112.loc[:,"位置"]!="A17-J")&\
            (df_112.loc[:,"位置"]!="LS205")&\
            (df_112.loc[:,"位置"]!="沪试剂/库存/沪试料/E栋丙类库2楼/E203中型货架库/YL-YP-HG")&\
            (~df_112.loc[:,"位置"].str.contains("112"))&\
            (~df_112.loc[:,"位置"].str.contains("退货位置"))&\
            (~df_112.loc[:,"位置"].str.contains("报废位置"))&\
            (~df_112.loc[:,"位置"].str.contains("质量管理"))&\
            (~df_112.loc[:,"项目号末尾"].str.contains("M"))&\
            (~df_112.loc[:,"位置"].str.contains("包材库"))


df_112_Done = df_112.loc[Condition_112,:]
df_112_Done = df_112_Done[["产品/内部参考",'不带C项目','位置','批次/序列号码','数量','单位',"项目号末尾"]]
#sumif 按照不带C项目  并merge到主表
grouped_112 = df_112_Done.groupby("不带C项目")["数量"].sum()
df_112_Done_with_sumif = pd.merge(df_112_Done,grouped_112,how="left",left_on="不带C项目",right_on="不带C项目")
#112_计划工作台用导出
df_112_Done_with_sumif.to_excel(Output_path_112_计划工作台用)
print("112_Done")


