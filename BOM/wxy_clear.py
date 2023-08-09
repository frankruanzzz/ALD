
import pandas as pd

info_product_path = f''
info_plan_platform_path = f""
# path of file
sheet_info = ""
sheet_plan = ""
#choose sheet
df_info = pd.read_excel(info_product_path,sheet_name=sheet_info)
df_plan = pd.read_excel(info_plan_platform_path,sheet_name=sheet_plan)
#read info

list_drop_info= []
#list contain the column which you want to delete in df_info
 
df_info.drop(list_drop_info,axis=1,inplace=True)
# execute drop

key = ""
#set key column to contact two df

df_info_wxy = pd.merge(df_info,df_plan,how="inner",on=key)
#merge

list_drop_plan =[]
#list contain the column which you want to delete 

df_info_wxy.drop(list_drop_plan,axis=1,inplace=True)
# execute drop

out_put= ""
df_info_wxy.to_excel(out_put)