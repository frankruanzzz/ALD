import pandas as pd
import numpy as np
path_New_bom = "/Users/frank/Downloads/ALD/BOM/New_bom_object.xlsx"
path_info_bom = '/Users/frank/Downloads/ALD/BOM/敏感性_for_test.xlsx'

df_info_bom = pd.read_excel(path_info_bom)
df_New_bom = pd.read_excel(path_New_bom)

s1 = df_New_bom.loc[:,"包装"]
df_New_bom.loc[:,"项目"] = s1.str.split("-").str[0]
df_New_bom.loc[:,"PKG"] = s1.str.split("-").str[1]

two_groups = '(\d+(?:\.\d+)?)([a-zA-Z|μ]+)'
df_letter_digit= df_New_bom.loc[:,"PKG"].str.extract(two_groups,expand=True)
df_New_bom.loc[:,"PKG_qty"] = df_letter_digit.loc[:,0]
df_New_bom.loc[:,"PKG_unit"] = df_letter_digit.loc[:,1]

info_done = pd.merge(df_New_bom,df_info_bom,how= 'inner',on= "项目")

info_done.drop(["储存温度","储存温度（内部）",'湿','热','光','气','敏感性'],axis=1,inplace=True)

info_done["tag_1"]= '1'
info_done["tag_2"]= ''
info_done.loc[:,"产品变体/内部参考"]= info_done["包装"]
info_done["产品/ID"]= ''
info_done["产品变体/ID"]= ''
info_done["数量"]= '1'
info_done["单位"]= '个'
info_done["工艺/ID"]= ''
info_done["BOM明细行/消耗在作业/ID"]= ''
info_done.loc[:,"BOM明细行/组件/内部参考"]= info_done["项目"]
info_done["BOM明细行/组件/ID"]= ''
info_done["BOM明细行/数量"]= ''
info_done.loc[:,"BOM明细行/组件/单位"]= info_done["计量单位"]
info_done.loc[:,"编号"]= info_done["包装"]
info_done.loc[:,"主原材料"]= info_done["项目"]
info_done["主原材料/外部 ID"]= ''
#print(info_done.head())
info_done.to_excel("/Users/frank/Downloads/ALD/BOM/BOM_info_auto_half_done.xlsx")
