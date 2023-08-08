import pandas as pd

info_product_path = '/Users/frank/Downloads/ALD/BOM/test_wxy.xlsx'

#后续替换成sheet[X] 用顺序来调用这个sheet
stock_111_path = "/Users/frank/Downloads/ALD/BOM/test_111.xlsx"

stock_112_path = "/Users/frank/Downloads/stock.quant_112_OK_Aug08.xlsx"

df_info = pd.read_excel(info_product_path)
df_111 = pd.read_excel(stock_111_path)
df_112 = pd.read_excel(stock_112_path)
#read info

#list_drop_info= ["项目号带C","备注综合","色谱溶剂","投入项类型","备注1","备注2","备注3","保质期","备注4","12年至今销量","2012至今销量（重量）",	"近2年销量重量","近1年销量重量","最小值重量","121库存",	121,"缺数1","多1","粤库存",131,"缺数2","多2","川库存",151,"缺数3","多3","华中库存",161,	"缺数4","多4","外仓缺数","缺数重量","出口-最小值","出口-返工","出口-wip","出口-库存","出口缺数","出口缺数重量","研发用量","生产wip",111,"外仓多库存","总库存数量","总库存重量",112,"差异数量","差异重量","对比","总库存重量/近2年重量",	"外仓库存+111",	"最大值","补到最大值个数","采购量","缺数重量+采购量+订单量","总库存重量/（最小值重量+缺数重量）","订单数量","订单总量","近一次收货时间","近一次收货数量","近一次收货单位","备注","单位","单位.1"]
#list contain the column which you want to delete in df_info
list_drop_112 = ["产品/内部参考","位置","批次/序列号码","单位"]
list_drop_111 =["不带C项目","位置","批次/序列号码","数量",'预留数量']
#df_info.drop(list_drop_info,axis=1,inplace=True)
df_112.set_index("不带C项目",inplace=True)
df_112.drop(list_drop_112,axis=1,inplace=True)
df_112.rename(columns={"数量":"112_数量"},inplace=True)

grouped = df_111.groupby("产品/内部参考")["可用数"].sum()

print(grouped)

# #set key column to contact two df
df_info_wxy = pd.merge(df_info,df_112,how="left",left_on="投入项",right_on="不带C项目")

# merge info with 112 on key 投入项 and 不带C项目
df_info_wxy_OK =pd.merge(df_info_wxy,grouped,how="left",left_on="内部参考",right_on="产品/内部参考")

out_put= "/Users/frank/Downloads/ALD/BOM/test_111_OK.xlsx"
df_info_wxy_OK.to_excel(out_put)
print("Done")