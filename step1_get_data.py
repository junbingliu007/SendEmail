import os
import shutil
import logging
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from common_utils import *


# 设置日志
logging.basicConfig(filename='jira_scraper.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

driver = login_jira()

# 清除缓存
gen_py_path = os.path.join(os.environ["LOCALAPPDATA"], "Temp", "gen_py")
if os.path.exists(gen_py_path):
    shutil.rmtree(gen_py_path)

# 读取 Excel
excel_file = get_excel_file_url()
sheet_name = "Sheet1"
df = pd.read_excel(excel_file, sheet_name=sheet_name, engine="openpyxl", header=0)
df.columns = ["RootCause", "Email", "Name","JiraStatus","Channel","Id", "JiraKey", "SendEmail","EmailStatus","SentDate","Remark"]


# 显式转换为字符串类型，避免 FutureWarning
df["RootCause"] = df["RootCause"].astype(str)
df["JiraStatus"] = df["JiraStatus"].astype(str)
df["Id"] = df["Id"].astype(str)
df["Email"] = df["Email"].astype(str)
df["Name"] = df["Name"].astype(str)



# 遍历每个 JiraKey
base_url_prefix = "https://jira.devops.nonprod.empf.local/jira/browse/"
for index, row in df.iterrows():
    issue_key = str(row["JiraKey"])
    if pd.isna(issue_key):
        continue

    url = base_url_prefix + issue_key
    driver.get(url)

    root_cause = ""
    jira_status = ""
    Id = ''
    reporter_email = ''
    reporter_name = ''

    try:
        cause_element = driver.find_element(By.ID, "customfield_11317-val")
        root_cause = cause_element.text or ""
        logging.info(f"{issue_key} - RootCause: {root_cause}")
    except Exception as e:
        logging.warning(f"{issue_key} - RootCause not found: {e}")

    try:
        status_element = driver.find_element(By.ID, "opsbar-transitions_more")
        jira_status = status_element.text or ""
        logging.info(f"{issue_key} - JiraStatus: {jira_status}")
    except Exception as e:
        logging.warning(f"{issue_key} - JiraStatus not found: {e}")

    try:
        ID_element = driver.find_element(By.ID,"customfield_11310-val")
        Id = ID_element.text
        logging.info(f"{issue_key} - ID: {Id}")
    except Exception as e:
        logging.warning(f"{issue_key} - ID not found: {e}")

    try:
        email_element = driver.find_element(By.CSS_SELECTOR, "#customfield_11303-val a")
        reporter_email = email_element.text or ""
        logging.info(f"{issue_key} - Reporter Email: {reporter_email}")
    except Exception as e:
        logging.warning(f"{issue_key} - Reporter Email not found: {e}")

    try:
        name_element = driver.find_element(By.ID, "customfield_11300-val")
        reporter_name = name_element.text or ""
        logging.info(f"{issue_key} - Reporter Name: {reporter_name}")
    except Exception as e:
        logging.warning(f"{issue_key} - Reporter Name not found: {e}")

    # 写回 DataFrame
    df.at[index, "RootCause"] = root_cause.strip()
    df.at[index, "JiraStatus"] = jira_status.strip()
    df.at[index, "Id"] = Id.strip()
    df.at[index, "Email"] = reporter_email.strip()
    df.at[index, "Name"] = reporter_name.strip()
# 保存 Excel
df.to_excel(excel_file, sheet_name=sheet_name, index=False, engine="openpyxl")
logging.info("Excel 文件已更新完成。")
driver.quit()