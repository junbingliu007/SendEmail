import os
import shutil
import logging
import pandas as pd
from selenium import webdriver
from common_utils import *
from selenium.webdriver.common.action_chains import ActionChains
from common_utils import *



# 设置日志
logging.basicConfig(filename='jira_scraper.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 初始化 WebDriver
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


filtered_df = df[
    (df["EmailStatus"] == "sent") &
    (~df["JiraStatus"].isin(["INSUFFICIENT INFO", "Proposed To Close"]))
]


# 提取所有非空的 JiraKey
issue_keys = filtered_df["JiraKey"].dropna().astype(str).tolist()



# 遍历每个 JiraKey
base_url_prefix = "https://jira.devops.nonprod.empf.local/jira/browse/"
for issue_key in issue_keys:

    if pd.isna(issue_key):
        continue

    url = base_url_prefix + issue_key
    driver.get(url)

    # 刷新当前页面
    # driver.refresh()

    try:
        status_element = wait_for_element(driver, By.ID, "opsbar-transitions_more",condition="clickable")
        # status_element = driver.find_element(By.ID, "opsbar-transitions_more")
        if status_element.text == "Proposed To Close":
            continue
        status_element.click()
        time.sleep(0.5)
        # 获取所有状态标签
        lozenges: list[WebElement] = []
        try:
            lozenges = wait_for_element(driver, By.CSS_SELECTOR, ".issueaction-workflow-transition .transition-label",condition="present")
        except Exception as e:
            logging.error(e)

        # lozenges = driver.find_elements(By.CSS_SELECTOR, ".issueaction-workflow-transition .transition-label")
        for lozenge in lozenges:
            # 显式等待 lozenge 中的第三个 div 出现
            third_div = WebDriverWait(lozenge, 10).until(
                lambda lz: lz.find_elements(By.TAG_NAME, "div")[2]
            )

            # 显式等待 span 元素出现并获取文本
            span_element = WebDriverWait(third_div, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "span"))
            )
            span_text = span_element.text

            if span_text == "PROPOSED TO CLOSE":
                try:
                    actions = ActionChains(driver)
                    actions.move_to_element(third_div).click().perform()
                    WebDriverWait(driver, 10).until(EC.staleness_of(third_div))  # 等待页面响应
                except Exception as e:
                    print(f"{issue_key} - Error clicking lozenge: {e}")
                    logging.warning(f"{issue_key} - Error clicking lozenge: {e}")
    except Exception as e:
        print(f"{issue_key} - JiraStatus element not found:  {e}")
        logging.warning(f"{issue_key} - JiraStatus element not found: {e}")








