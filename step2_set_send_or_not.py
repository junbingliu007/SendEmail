import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from common_utils import *

# Excel 文件路径和读取设置
excel_file = get_excel_file_url()
sheet_name = "Sheet1"

# 读取 Excel
df = pd.read_excel(excel_file, sheet_name=sheet_name, engine="openpyxl", header=0)
df.columns = ["RootCause", "Email", "Name", "JiraStatus", "Channel", "Id", "JiraKey", "SendEmail", "EmailStatus", "SentDate", "Remark"]

df["Remark"] = df["Remark"].astype(str)
df["SendEmail"] = df["SendEmail"].astype(str)

driver = login_jira()

# 遍历每个 JiraKey
base_url_prefix = "https://jira.devops.nonprod.empf.local/jira/browse/"
for index, row in df.iterrows():
    issue_key = str(row["JiraKey"])
    if pd.isna(issue_key):
        continue

    url = base_url_prefix + issue_key
    driver.get(url)

    # 查找元素并截图
    element = driver.find_element(By.ID, "issue_actions_container")

    # 使用 JavaScript 滚动元素到可视区域
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)

    try:
        box_content = confirm_box(driver)

        choice = box_content['choice']
        remark = box_content['remark']
        if choice == 'Yes':
            df.at[index, "SendEmail"] = "yes"
            df.at[index, "Remark"] = ""
        elif choice == 'No':
            df.at[index, "SendEmail"] = "no"
            df.at[index, "Remark"] = remark
        df.to_excel(excel_file, sheet_name=sheet_name, index=False, engine="openpyxl")

    except Exception as e:
        print(f"处理 {url} 时发生错误：{e}")
        continue




