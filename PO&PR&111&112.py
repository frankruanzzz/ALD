import pandas as pd
import numpy as np

date = input("输入日期_")

path_Po = f"F:\RYH\物料相关\每日数据源\{date}\purchase.order.xlsx"
Output_path_Po =f"F:\RYH\物料相关\每日数据源\{date}\purchase.order_OK_{date}.xlsx"

df_Po = pd.read_excel(path_Po)
df_Po.loc[:,"不带C项目"]= df_Po["订单行/产品/内部参考"].str[:7]
df_Po.loc[:,"订单行/数量"]= df_Po.loc[:,"订单行/数量"].astype("float")
df_Po.loc[:,"订单行/已接收数量"]= df_Po.loc[:,"订单行/已接收数量"].astype("float")
df_Po.loc[:,"未到货数"] = (df_Po.loc[:,"订单行/数量"])-(df_Po.loc[:,"订单行/已接收数量"])
df_Po.loc[:,"未到货数"]=df_Po.loc[:,"未到货数"].astype("float")
df_Po.loc[:,"未到货数占比"] = (df_Po.loc[:,"未到货数"])/(df_Po.loc[:,"订单行/数量"])
df_Po.loc[:,"未到货数占比"]=df_Po.loc[:,"未到货数占比"].astype("float")
df_Po.loc[:,"订单关联"] = df_Po["订单关联"].fillna(method="ffill")
df_Po.loc[:,"单据日期"] = df_Po["单据日期"].fillna(method="ffill")
df_Po.loc[:,"采购员"] = df_Po["采购员"].fillna(method="ffill")
df_Po.loc[:,"跟单员"] = df_Po["跟单员"].fillna(method="ffill")
df_Po.loc[:,"已结单标志位"] = df_Po["已结单标志位"].fillna(method="ffill")
df_Po.loc[:,"源单据"] = df_Po["源单据"].fillna(method="ffill")
df_Po.loc[:,"状态"] = df_Po["状态"].fillna(method="ffill")
df_Po.loc[:,"标签"] = df_Po["标签"].fillna("0")
def Condition_C (df):
    return (df.loc[:,"未到货数"] >0)&\
           (df.loc[:,"已结单标志位"]==False)
def condition_112(df):
    return  ~df["标签"].str.contains("重新备货|待退款|待结算")          
i =df_Po.loc[Condition_C,:]
i= i.loc[condition_112,:]
i = i[["订单关联","单据日期","订单行/产品/内部参考","不带C项目","订单行/数量","订单行/已接收数量","未到货数","订单行/单位","采购员","跟单员",'已结单标志位','标签','源单据','状态','采购申请标签']]
i.to_excel(Output_path_Po)
print("Po_Done")

path_112 = f"F:\RYH\物料相关\每日数据源\{date}\stock.quant_112.xlsx"
Output_path_112 = f"F:\RYH\物料相关\每日数据源\{date}\stock.quant_112_OK_{date}.xlsx"
df_112 = pd.read_excel(path_112)
df_112.loc[:,"不带C项目"]= df_112["产品/内部参考"].str[:7]
df_112.loc[:,"数量"]= df_112.loc[:,"数量"].astype("float")
df_112.set_index("产品/内部参考",inplace=True)
def Condition_112(df_112):
    return (df_112.loc[:,"位置"]!="BL112")&\
            (df_112.loc[:,"位置"]!="LT112")&\
            (df_112.loc[:,"位置"]!="FC112")&\
            (df_112.loc[:,"位置"]!="YC112")&\
            (df_112.loc[:,"位置"]!="ZC112")&\
            (df_112.loc[:,"位置"]!="YP-DJ")&\
            (df_112.loc[:,"位置"]!="YP-DJ")&\
            (df_112.loc[:,"位置"]!="YP-HG")&\
            (df_112.loc[:,"位置"]!="YP-HG2-8")&\
            (df_112.loc[:,"位置"]!="BLYF")&\
            (df_112.loc[:,"位置"]!="A17-J")&\
            (df_112.loc[:,"位置"]!="YF112")&\
            (df_112.loc[:,"位置"]!="LS205")&\
            (df_112.loc[:,"位置"]!="不良仓/库存/沪试料/E栋丙类库/报废位置/YL-BL112")&\
            (df_112.loc[:,"位置"]!="不良仓/库存/沪试料/E栋丙类库/报废位置/YL-FC112")&\
            (df_112.loc[:,"位置"]!="沪试剂/库存/沪试料/E栋丙类库2楼/E203中型货架库/YL-YP-HG")&\
            (df_112.loc[:,"位置"]!="不良仓/库存/沪试料/E栋丙类库/报废位置/YL-LT112")&\
            (df_112.loc[:,"位置"]!="不良仓/库存/沪试料/E栋丙类库/报废位置/YL-YC112")&\
            (df_112.loc[:,"位置"]!="沪试剂/库存/沪试料/E栋丙类库2楼/YL-BY112")&\
            (df_112.loc[:,"位置"]!="沪试剂/质量管理/沪试料/YL-LS205")&\
            (df_112.loc[:,"位置"]!="沪试剂/质量管理/沪试料/YL-YP-DJ")&\
            (df_112.loc[:,"位置"]!="不良仓/库存/沪试料/E栋丙类库/报废位置/YL-FC112/YP-HG")
df_112_Done = df_112.loc[Condition_112,:]
df_112_Done = df_112_Done[['不带C项目','位置','批次/序列号码','数量','单位']]
df_112_Done.to_excel(Output_path_112)
print("112_Done")

path_111 = f"F:\RYH\物料相关\每日数据源\{date}\stock.quant_111.xlsx"
Output_path_111 = f"F:\RYH\物料相关\每日数据源\{date}\stock.quant_111_OK_{date}.xlsx"
df_111 = pd.read_excel(path_111)
df_111.loc[:,"不带C项目"]= df_111["产品/内部参考"].str[:7]
df_111.loc[:,"数量"]= df_111.loc[:,"数量"].astype("float")
df_111.loc[:,"预留数量"]= df_111.loc[:,"预留数量"].astype("float")
df_111.loc[:,"可用数"] = (df_111.loc[:,"数量"])-(df_111.loc[:,"预留数量"])
df_111.set_index("产品/内部参考",inplace=True)
def Condition_111(df):
    return (df.loc[:,"位置"]!="FC111")&\
            (df.loc[:,"位置"]!="LS111")&\
            (df.loc[:,"位置"]!="WL111")&\
            (df.loc[:,"位置"]!="ZJ111")&\
            (df.loc[:,"位置"]!="PICKTO")&\
            (df.loc[:,"位置"]!="yf111")&\
            (df.loc[:,"位置"]!="gy111")&\
            (df.loc[:,"位置"]!="不良仓/库存/沪试成报废区/报废位置/FC111")&\
            (df.loc[:,"位置"]!="不良仓/库存/沪试成报废区/报废位置/FJ111")&\
            (df.loc[:,"位置"]!="不良仓/库存/沪试成报废区/报废位置/LS111")&\
            (df.loc[:,"位置"]!="不良仓/库存/沪试成报废区/报废位置/U98")&\
            (df.loc[:,"位置"]!="不良仓/库存/沪试成报废区/报废位置/WL111")&\
            (df.loc[:,"位置"]!="不良仓/库存/沪试成报废区/报废位置/XJ111")&\
            (df.loc[:,"位置"]!="不良仓/库存/沪试成报废区/报废位置/ZJ111")&\
            (df.loc[:,"位置"]!="沪试剂/库存/沪试成/退货区/退货位置/TH111")&\
            (df.loc[:,"位置"]!="不良仓/库存/沪试成报废区/报废位置/FW111")&\
            (df.loc[:,"位置"]!="不良仓/库存/沪试成报废区/报废位置/SH-HB111")&\
            (df.loc[:,"位置"]!="不良仓/库存/沪试成报废区/报废位置/SH-HB-US")&\
            (df.loc[:,"可用数"] >=1)         
df_111_Done = df_111.loc[Condition_111,:]
df_111_Done= df_111_Done[["不带C项目","位置","批次/序列号码",'数量','预留数量','可用数']]
df_111_Done.to_excel(Output_path_111)
print("111_Done")

path_pr =f"F:\RYH\物料相关\每日数据源\{date}\purchase.requisition.xlsx"
Output_path_Pr =f"F:\RYH\物料相关\每日数据源\{date}\purchase.requisition_OK_{date}.xlsx"
df_pr = pd.read_excel(path_pr)
df_pr.loc[:,"不带C项目"]= df_pr["采购的产品/产品/内部参考"].str[:7]
df_pr.set_index("采购的产品/产品/内部参考",inplace=True)
df_pr= df_pr[['编号','订购日期','申请截止日期','不带C项目','采购的产品/数量','采购的产品/产品计量单位','采购申请标签','采购员','源单据','状态','操作类型','说明']]
df_pr.to_excel(Output_path_Pr)
print("Pr_Done")
print("All_DOne")


