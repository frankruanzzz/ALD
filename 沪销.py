import pandas as pd
import numpy as np
path_stock ="/Users/frank/Downloads/stock.picking（沪销）.xlsx"
path_sale_order_line= "/Users/frank/Downloads/sale.order.line11.3.xlsx"


df_stock = pd.read_excel(path_stock)
df_saleorderline = pd.read_excel(path_sale_order_line)

df_stock.loc[:,"创建时间"] = df_stock["创建时间"].fillna(method="ffill")
df_stock.loc[:,"销售订单/客户/责任客服"] = df_stock["销售订单/客户/责任客服"].fillna(method="ffill")
df_stock.loc[:,"安排的日期"] = df_stock["安排的日期"].fillna(method="ffill")
df_stock.loc[:,"源文档"] = df_stock["源文档"].fillna(method="ffill")
df_stock.loc[:,"源位置"] = df_stock["源位置"].fillna(method="ffill")
df_stock.loc[:,"差"]= df_stock.loc[:,"库存移动不在包裹里/初始需求"]-df_stock.loc[:,"库存移动不在包裹里/已预留数量"]-df_stock.loc[:,"库存移动不在包裹里/完成数量"]

def condition_源文档_退回订单剔除条件(df):
    return  ~df["源文档"].str.contains("退回")

df_stock= df_stock.loc[condition_源文档_退回订单剔除条件]
df_stock.loc[:,"安排的日期"]= df_stock["源文档"] + df_stock["库存移动不在包裹里/产品/内部参考"]

df_saleorderline.loc[:,"Index"] = df_saleorderline["订单关联"] + df_saleorderline["产品/内部参考"]



沪销_Done = pd.merge(df_stock,df_saleorderline,how= 'inner',left_on="安排的日期",right_on="Index")


沪销_Done.to_excel('/Users/frank/Downloads/stock.picking_test01.xlsx')
print("done")