import time

import pandas as pd
import webbrowser
import shutil
import os
import subprocess

from selenium import webdriver
from selenium.webdriver.common.by import By
from common_utils import *

# 清除缓存（可选）
gen_py_path = os.path.join(os.environ["LOCALAPPDATA"], "Temp", "gen_py")
if os.path.exists(gen_py_path):
    shutil.rmtree(gen_py_path)

# Excel 文件路径和读取设置
excel_file = get_excel_file_url()
sheet_name = "Sheet1"

# 读取整个 Excel 表格（第一行是表头）
df = pd.read_excel(excel_file, sheet_name=sheet_name, engine="openpyxl", header=0)

# 设置正确的列名顺序
df.columns = ["RootCause", "Email", "Name","JiraStatus","Channel","Id","JiraKey", "SendEmail","EmailStatus","SentDate","Remark"]

filtered_df = df[(df["EmailStatus"] == "sent") & (df["JiraStatus"] != "INSUFFICIENT INFO")]


# 提取所有非空的 JiraKey
issue_keys = filtered_df["JiraKey"].dropna().astype(str).tolist()

# 构造所有链接
base_url_prefix = "https://jira.devops.nonprod.empf.local/jira/browse/"
jira_urls = [base_url_prefix + issue_key for issue_key in issue_keys]


# 每批最多打开 50 个链接
batch_size = 50

for i in range(0, len(jira_urls), batch_size):
    batch_urls = jira_urls[i:i + batch_size]

    # 启动浏览器并登录
    driver = login_jira()

    # 打开当前批次的链接（新标签页）
    for url in batch_urls:
        driver.execute_script(f"window.open('{url}', '_blank');")
        time.sleep(0.5)

    input(f"\n已打开第 {i//batch_size + 1} 批链接，请关闭浏览器后按 Enter 键继续...")

    # 关闭当前 WebDriver 实例
    driver.quit()


