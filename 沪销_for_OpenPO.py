import pandas as pd
import numpy as np
path_stock =  "F:/RYH/Open_PO_更新/stock.picking (沪销).xlsx"

path_sale_order_line= "F:/RYH/Open_PO_更新/sale.order.line.xlsx"

df_stock = pd.read_excel(path_stock)
df_saleorderline = pd.read_excel(path_sale_order_line)

df_stock.loc[:,"创建时间"] = df_stock["创建时间"].fillna(method="ffill")
df_stock.loc[:,"销售订单/客户/责任客服"] = df_stock["销售订单/客户/责任客服"].fillna(method="ffill")
df_stock.loc[:,"安排的日期"] = df_stock["安排的日期"].fillna(method="ffill")
df_stock.loc[:,"联系人"] = df_stock["联系人"].fillna(method="ffill")
df_stock.loc[:,"源文档"] = df_stock["源文档"].fillna(method="ffill")
df_stock.loc[:,"源位置"] = df_stock["源位置"].fillna(method="ffill")
df_stock.loc[:,"差"]= df_stock.loc[:,"库存移动不在包裹里/初始需求"]-df_stock.loc[:,"库存移动不在包裹里/已预留数量"]-df_stock.loc[:,"库存移动不在包裹里/完成数量"]
def condition_源文档_退回订单剔除条件(df):
    return  ~df["源文档"].str.contains("退回")

df_stock= df_stock.loc[condition_源文档_退回订单剔除条件]
# 执行 沪销中的筛选条件 （剔除 退回的订单）
df_stock.loc[:,"安排的日期"]= df_stock["源文档"] + df_stock["库存移动不在包裹里/产品/内部参考"]
# 沪销中 创建 Index
df_saleorderline.loc[:,"Index"] = df_saleorderline["订单关联"] + df_saleorderline["产品/内部参考"]
# 订单明细行中创建 Index
grouped_saleorder=  df_saleorderline.groupby("Index")["小计"].sum()

# sumif index & 小计

selected_columns = ["Index","订单关联/条款和条件"]
# 选取需要的字段
new_df = df_saleorderline[selected_columns]
# 创建对用字段的df

df_saleorderline_OK = pd.merge(new_df,grouped_saleorder,how="left",on="Index")
# index sumif小计OK & index 备注OK

沪销_Done = pd.merge(df_stock,df_saleorderline_OK,how= 'left',left_on="安排的日期",right_on="Index")
# 沪销 & index 小计&备注 
沪销_Done = 沪销_Done.drop_duplicates()
沪销_Done = 沪销_Done[["创建时间","销售订单/客户/责任客服","联系人","安排的日期","源文档","库存移动不在包裹里/产品/内部参考","源位置","库存移动不在包裹里/初始需求","库存移动不在包裹里/已预留数量","库存移动不在包裹里/完成数量","追踪参考","差","Index","小计","订单关联/条款和条件"]]
沪销_Done.to_excel("F:/RYH/Open_PO_更新/沪销_OK.xlsx")
print("Done")