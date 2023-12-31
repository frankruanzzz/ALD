from selenium import webdriver
from bs4 import BeautifulSoup
import json
import time
import pandas as pd

path_driver = 'D:\ALD\Chromedriver\chromedriver-win64\chromedriver.exe'
driver = webdriver.Chrome()

path = 'D:\ALD\毕得&阿拉丁Selenium数据汇总CAS总表.xlsx'

# sheet 文件名
df = pd.read_excel(path,sheet_name='falied_ALL')
# sheet 文件名

cas_list = df['CAS']

# JSON 文件名
output_filename = 'output_data_failed_All_P2.json'  
# JSON 文件名
try:
    with open(output_filename, 'r', encoding='utf-8') as json_file:
        output_data = json.load(json_file) 
except FileNotFoundError:
    output_data = []

for i in cas_list:
    try:
        driver.get(f"https://www.bidepharm.com/products/{i}.html")
        time.sleep(1)
        driver.implicitly_wait(25)
        time.sleep(2)
        page_source = driver.page_source
        soup = BeautifulSoup(page_source, "html.parser")
        cas_block = soup.find("div", class_="products-title")
        cas_number_OK = cas_block.find("h1").text.strip()

        product_block = soup.find("div", class_="products-top-kuang")
        product_id_OK = product_block.find("b", id='first_bd').text.strip()

        table = soup.find('div', class_='products-top-table products-table')
        rows = table.find_all('tr')
        all_cell_text = []
        table_data = [row.find_all(['tr', 'td']) for row in rows]

        for row_data in table_data:
            cell_text = [cell.text for cell in row_data]
            all_cell_text.append(cell_text)

        cleaned_data = [[cell_for_clear.strip() for cell_for_clear in row_clear] for row_clear in all_cell_text]
        df_sorted = pd.DataFrame(cleaned_data[1:], columns=cleaned_data[0])

        # 添加当前循环的数据到列表
        data = {
            "CAS_Number": cas_number_OK,
            "BD_Product_ID": product_id_OK,
            "Data": df_sorted.to_dict(orient='records')
        }
        output_data.append(data)
        
        # 将当前循环的数据写入 'output_data' JSON 文件
        with open(output_filename, 'w', encoding='utf-8') as json_file:
            json.dump(output_data, json_file, ensure_ascii=False, indent=4)
    
    except Exception as e:
        print(f"Error while processing {i}: {e}")

# 关闭 WebDriver
driver.quit()

print(f"数据已存储到 {output_filename} 文件中")