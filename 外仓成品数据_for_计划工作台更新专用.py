import pandas as pd
import numpy as np

date = input("输入日期_")



def Condition_drop_info(df_112):
    return (~df_112.loc[:,"位置"].str.contains("报废位置"))&\
           (~df_112.loc[:,"位置"].str.contains("出货"))
    
            # (df_112.loc[:,"位置"]!="不良仓/库存/鄂成报废区/报废位置/WH-FC161")&\
            # (df_112.loc[:,"位置"]!="不良仓/库存/鄂成报废区/报废位置/WH-U98")&\
            # (df_112.loc[:,"位置"]!='不良仓/库存/鄂成报废区/报废位置/WH-HB161')&\
            # (df_112.loc[:,"位置"]!='不良仓/库存/鄂成报废区/报废位置/WH-FJ161')&\
            # (df_112.loc[:,"位置"]!='不良仓/库存/鄂成报废区/报废位置/WH-LS161')&\
            # (df_112.loc[:,"位置"]!="鄂成/出货")&\
            #     (df_112.loc[:,"位置"]!="不良仓/库存/津成报废区/报废位置/LS121")&\
            #     (df_112.loc[:,"位置"]!="不良仓/库存/津成报废区/报废位置/TJ-FC121")&\
            #     (df_112.loc[:,"位置"]!="不良仓/库存/津成报废区/报废位置/TJ-XJ121")&\
            #     (df_112.loc[:,"位置"]!="不良仓/库存/津成报废区/报废位置/TJ-FJ121")&\
            #     (df_112.loc[:,"位置"]!="不良仓/库存/津成报废区/报废位置/TJ-HB121")&\
            #     (df_112.loc[:,"位置"]!="津成/出货")&\
            #         (df_112.loc[:,"位置"]!="不良仓/库存/粤成报废区/GD-LS")&\
            #         (df_112.loc[:,"位置"]!="不良仓/库存/粤成报废区/GD-U98")&\
            #         (df_112.loc[:,"位置"]!="不良仓/库存/粤成报废区/报废位置")&\
            #         (df_112.loc[:,"位置"]!="不良仓/库存/粤成报废区/报废位置/GD-FC131")&\
            #         (df_112.loc[:,"位置"]!="不良仓/库存/粤成报废区/报废位置/GD-FJ131")&\
            #         (df_112.loc[:,"位置"]!="不良仓/库存/粤成报废区/报废位置/GD-HB131")&\
            #         (df_112.loc[:,"位置"]!="粤成/出货")&\
            #             (df_112.loc[:,"位置"]!="不良仓/库存/川成报废区/CD-FC151")&\
            #             (df_112.loc[:,"位置"]!="不良仓/库存/川成报废区/CD-U98")&\
            #             (df_112.loc[:,"位置"]!='不良仓/库存/川成报废区/CD-FJ151')&\
            #             (df_112.loc[:,"位置"]!='不良仓/库存/川成报废区/CD-HB151')&\
            #             (df_112.loc[:,"位置"]!="川成/出货")
           
def Condition_qty (df):
    return df.loc[:,"可用数"] >=0
# 分拨仓数据 提取 
filepath_分拨仓 = f"F:\RYH\物料相关\每日数据源\{date}\分拨仓.xlsx" 
df_分拨仓 = pd.read_excel(filepath_分拨仓)
df_CD =df_分拨仓[df_分拨仓["位置"].str.contains("CD")]
df_GD =df_分拨仓[df_分拨仓["位置"].str.contains("GD")]
df_WH =df_分拨仓[df_分拨仓["位置"].str.contains("WH")]
df_TJ =df_分拨仓[df_分拨仓["位置"].str.contains("TJ")]


filepath_川成 = f"F:\RYH\物料相关\每日数据源\{date}\川成.xlsx"
Output_path_川成 =f"F:\RYH\物料相关\每日数据源\{date}\川成_{date}_计划工作台用.xlsx"

df_川成 = pd.read_excel(filepath_川成)
#插入 contact
df_川成 = pd.concat([df_川成,df_CD],axis=0)
#-----
df_川成.loc[:,"不带C项目"]= df_川成["产品/内部参考"].str[:7]
df_川成.loc[:,"数量"]= df_川成.loc[:,"数量"].astype("float")
df_川成.loc[:,"预留数量"]= df_川成.loc[:,"预留数量"].astype("float")
df_川成.loc[:,"可用数"] = (df_川成.loc[:,"数量"])-(df_川成.loc[:,"预留数量"])
df_川成.set_index("产品/内部参考",inplace=True)

i_川成 = df_川成.loc[Condition_drop_info,:]
川成_Done =i_川成.loc[Condition_qty,:]

#-------------
grouped_川成 = 川成_Done.groupby("产品/内部参考")["可用数"].sum()
# grouped_川成.rename(colums={"可用数_y":"SUMIF_可用数"},inplace=True)
川成_Done_with_sumif = pd.merge(川成_Done,grouped_川成,how="left",left_on="产品/内部参考",right_on="产品/内部参考")
川成_Done_with_sumif.to_excel(Output_path_川成)
#-------------

print("川成_Done")


filepath_鄂成 = f"F:\RYH\物料相关\每日数据源\{date}\鄂成.xlsx"
Output_path_鄂成 =f"F:\RYH\物料相关\每日数据源\{date}\鄂成_{date}_计划工作台用.xlsx"

df_鄂成 = pd.read_excel(filepath_鄂成)
#插入 contact
df_鄂成 = pd.concat([df_鄂成,df_WH],axis=0)
#-----

df_鄂成.loc[:,"不带C项目"]= df_鄂成["产品/内部参考"].str[:7]
df_鄂成.loc[:,"数量"]= df_鄂成.loc[:,"数量"].astype("float")
df_鄂成.loc[:,"预留数量"]= df_鄂成.loc[:,"预留数量"].astype("float")
df_鄂成.loc[:,"可用数"] = (df_鄂成.loc[:,"数量"])-(df_鄂成.loc[:,"预留数量"])
df_鄂成.set_index("产品/内部参考",inplace=True)        
i_鄂成 = df_鄂成.loc[Condition_drop_info,:]
鄂成_Done =i_鄂成.loc[Condition_qty,:]

grouped_鄂成 = 鄂成_Done.groupby("产品/内部参考")["可用数"].sum()
# grouped_川成.rename(colums={"可用数_y":"SUMIF_可用数"},inplace=True)
鄂成_Done_with_sumif = pd.merge(鄂成_Done,grouped_鄂成,how="left",left_on="产品/内部参考",right_on="产品/内部参考")
鄂成_Done_with_sumif.to_excel(Output_path_鄂成)


print("鄂成_Done")


filepath_津成 = f"F:\RYH\物料相关\每日数据源\{date}\津成.xlsx"
Output_path_津成 =f"F:\RYH\物料相关\每日数据源\{date}\津成_{date}_计划工作台用.xlsx"

df_津成 = pd.read_excel(filepath_津成)
#插入 contact
df_津成 = pd.concat([df_津成,df_TJ],axis=0)
#-----


df_津成.loc[:,"不带C项目"]= df_津成["产品/内部参考"].str[:7]
df_津成.loc[:,"数量"]= df_津成.loc[:,"数量"].astype("float")
df_津成.loc[:,"预留数量"]= df_津成.loc[:,"预留数量"].astype("float")
df_津成.loc[:,"可用数"] = (df_津成.loc[:,"数量"])-(df_津成.loc[:,"预留数量"])
df_津成.set_index("产品/内部参考",inplace=True)          
i_津成 = df_津成.loc[Condition_drop_info,:]
津成_Done =i_津成.loc[Condition_qty,:]

grouped_津成 = 津成_Done.groupby("产品/内部参考")["可用数"].sum()
# grouped_川成.rename(colums={"可用数_y":"SUMIF_可用数"},inplace=True)
津成_Done_with_sumif = pd.merge(津成_Done,grouped_津成,how="left",left_on="产品/内部参考",right_on="产品/内部参考")
津成_Done_with_sumif.to_excel(Output_path_津成)

print("津成_Done")


filepath_粤成 = f"F:\RYH\物料相关\每日数据源\{date}\粤成.xlsx"
Output_path_粤成 =f"F:\RYH\物料相关\每日数据源\{date}\粤成_{date}_计划工作台用.xlsx"

df_粤成 = pd.read_excel(filepath_粤成)
#插入 contact
df_粤成 = pd.concat([df_粤成,df_GD],axis=0)
#-----


df_粤成.loc[:,"不带C项目"]= df_粤成["产品/内部参考"].str[:7]
df_粤成.loc[:,"数量"]= df_粤成.loc[:,"数量"].astype("float")
df_粤成.loc[:,"预留数量"]= df_粤成.loc[:,"预留数量"].astype("float")
df_粤成.loc[:,"可用数"] = (df_粤成.loc[:,"数量"])-(df_粤成.loc[:,"预留数量"])
df_粤成.set_index("产品/内部参考",inplace=True)         
i_粤成 = df_粤成.loc[Condition_drop_info,:]
粤成_Done =i_粤成.loc[Condition_qty,:]

grouped_粤成 = 粤成_Done.groupby("产品/内部参考")["可用数"].sum()
# grouped_川成.rename(colums={"可用数_y":"SUMIF_可用数"},inplace=True)
粤成_Done_with_sumif = pd.merge(粤成_Done,grouped_粤成,how="left",left_on="产品/内部参考",right_on="产品/内部参考")
粤成_Done_with_sumif.to_excel(Output_path_粤成)
print("粤成_Done")


print("ALL_Done")