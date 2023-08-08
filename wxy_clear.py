import pandas as pd

info_product_path = input("输入产品信息表路径_")
sheet_info = '产品资料8.4'
info_plan_platform_path = input("输入计划工作台路径_")
sheet_plan = "计划工作台"
#choose sheet
df_info = pd.read_excel(info_product_path,sheet_name=sheet_info)
df_plan = pd.read_excel(info_plan_platform_path,sheet_name=sheet_plan)
#read info

list_drop_info= ["项目号带C","备注综合","色谱溶剂","投入项类型","备注1","备注2","备注3","保质期","备注4","12年至今销量","2012至今销量（重量）",	"近2年销量重量","近1年销量重量","最小值重量","121库存",	121,"缺数1","多1","粤库存",131,"缺数2","多2","川库存",151,"缺数3","多3","华中库存",161,	"缺数4","多4","外仓缺数","缺数重量","出口-最小值","出口-返工","出口-wip","出口-库存","出口缺数","出口缺数重量","研发用量","生产wip",111,"外仓多库存","总库存数量","总库存重量",112,"差异数量","差异重量","对比","总库存重量/近2年重量",	"外仓库存+111",	"最大值","补到最大值个数","采购量","缺数重量+采购量+订单量","总库存重量/（最小值重量+缺数重量）","订单数量","订单总量","近一次收货时间","近一次收货数量","近一次收货单位","备注","单位","单位.1"]
#list contain the column which you want to delete in df_info
 
df_info.drop(list_drop_info,axis=1,inplace=True)
# execute drop

list_drop_plan =["项目号","Is Published","Cas编号","密度","名称","规格或纯度","2年销\n重量","ProductVariantAvailable","大宗备货",	"包装数",	"包装量","参考单价","是否大包装/出口","细分类","保质期","2002至今销售数量",	"2年销数","津成成品库存","华南仓成品库存","西南仓成品库存",	"鄂成品库存",	"美国仓库存","津成半年计划量","华南仓半年计划量","西南仓半年计划量","武汉仓半年计划量","美国仓计划量","外仓缺货量","类型","投入项","工单号","aaa","原料复检","存货单位","贮存","沸点","熔点","安全库存+分仓半年计划量","openpo","沪试成+津成+粤成+川成","分装数量","分装重量","最终库存","产品类别","ABC","管控信息","套件"]
#list contain the column which you want to delete 
df_plan.drop(list_drop_plan,axis=1,inplace=True)
# execute drop


key = "内部参考"
# #set key column to contact two df
df_info_wxy = pd.merge(df_info,df_plan,how="left",on=key)
# #merge

#print(df_info_wxy.head(3))

out_put= "F:\RYH\物料相关\wxy-产品资料.xlsx"
df_info_wxy.to_excel(out_put)
print("Done")
