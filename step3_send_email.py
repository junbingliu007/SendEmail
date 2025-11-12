import subprocess
import time

import win32com.client
import pandas as pd
import shutil
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.action_chains import ActionChains
import logging
from common_utils import *

# 设置日志
logging.basicConfig(filename='jira_scraper.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 清除缓存（可选）
# gen_py_path = os.path.join(os.environ["LOCALAPPDATA"], "Temp", "gen_py")
# if os.path.exists(gen_py_path):
#     shutil.rmtree(gen_py_path)

# 初始化 WebDriver
driver = login_jira()

# Excel 文件路径和读取设置
excel_file = get_excel_file_url()
sheet_name = "Sheet1"

# 读取整个 Excel 表格
df = pd.read_excel(excel_file, sheet_name=sheet_name, engine="openpyxl", header=0)
df.columns = ["RootCause", "Email", "Name","JiraStatus","Channel","Id", "JiraKey", "SendEmail","EmailStatus","SentDate","Remark"]

# 显式转换为字符串类型，避免 FutureWarning
df["RootCause"] = df["RootCause"].astype(str)
df["JiraStatus"] = df["JiraStatus"].astype(str)
df["Id"] = df["Id"].astype(str)
df["Email"] = df["Email"].astype(str)
df["Name"] = df["Name"].astype(str)
df["SendEmail"] = df["SendEmail"].astype(str)
df["EmailStatus"] = df["EmailStatus"].astype(str)
df["SentDate"] = df["SentDate"].astype(str)
df["Remark"] = df["Remark"].astype(str)

# 设置起始和结束行索引


# 启动 Outlook 应用
outlook = win32com.client.Dispatch("Outlook.Application")
chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"


# 先筛选出需要发送邮件的行
filtered_df = df[(df["SendEmail"] != "no") & (df["EmailStatus"] != "sent")]
base_url_prefix = "https://jira.devops.nonprod.empf.local/jira/browse/"


# 遍历每一行
for index, row in filtered_df.iterrows():
    recipient_address = row["Email"]
    recipient_name = row["Name"]
    incident_num = row["Id"]
    prod_num = row["JiraKey"]
    root_cause = row["RootCause"] or "Other Data Issue"
    jira_status = row["JiraStatus"]
    channel = row["Channel"]

    if channel == "Admin Office" or channel == "Outreach_c" or channel == "Internal User":

        # 创建邮件
        mail = outlook.CreateItem(0)
        if channel == "Admin Office" or channel == "Outreach_c":
            mail.To = f"{recipient_address}; ticket@ifastepension.com"
        elif channel == "Internal User":
            mail.To = f"{recipient_address}"

        mail.CC = (
            "Stanislaus.SC.Tsang@pccw.com; "
            "Leo.CL.Chan@pccw.com; "
            "Yiki.YK.Kwan@pccw.com; "
            "bst@ifastepension.com; "
            "Brian.CM.Kwok@pccw.com; "
            "im@support.empf.org.hk; "
            "May.XM.He@pccw.com; "
            "Michael.L.Fan@pccw.com; "
            "Tobia.Feng@pccw.com"
        )

        # 设置主题和正文
        mail.Subject = f"Ticket Resolved – {incident_num}, {prod_num}"
        # 判断 JiraStatus
        if jira_status == "INSUFFICIENT INFO":
            mail.HTMLBody = f"""
                            <html>
                            <body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
                                <p>Dear {recipient_name},</p>

                                <p>We acknowledge receipt of the captioned incident.<br>
                                However, the information provided concerning this incident is insufficient for us to fully address the issue.
                                To enable us to assist you further, appreciate you could provide additional details about the incident.</p>

                                <p>此处有图片</p>

                                <p>For detail, please refer to the below Jira link:<br>
                                https://jira.devops.nonprod.empf.local/jira/browse/{prod_num}
                                    https://jira.devops.nonprod.empf.local/jira/browse/{prod_num}
                                </a></p>

                                <p>Once we receive the necessary information, we will be able to look into this matter and provide an appropriate response.</p>

                                <p>If we do not receive any feedback from you regarding this additional information within the next 2 days, we will consider the ticket as closed.</p>

                                <p>Best regards,<br>
                                Jun-Bing, Liu</p>
                            </body>
                            </html>
                            """
            embed_element_screenshot_in_email(driver, mail)
            mail.Display()

        elif jira_status == "Resolved" or jira_status == "Rejected" or jira_status == "Proposed To Close" or jira_status == "RELEASED":
            url = base_url_prefix = "https://jira.devops.nonprod.empf.local/jira/browse/" + prod_num
            driver.get(url)

            osc_ticket: list[WebElement] = driver.find_elements(By.CSS_SELECTOR, ".links-list .link-content a[data-issue-key^='OSC-']")

            if osc_ticket:
                if len(osc_ticket) > 5:
                    element = driver.find_element(By.ID, "show-more-links-link")
                    # 滚动到该元素
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                    time.sleep(0.5)
                    # 点击元素
                    ActionChains(driver).move_to_element(element).click().perform()
                osc_links = ""
                for osc in osc_ticket:
                    osc_links += f"https://jira.devops.nonprod.empf.local/jira/browse/{osc.text}\n"

                mail.Subject = f"Ticket Closed – {incident_num}, {prod_num}"

                mail.Body = (
                    f"Dear {recipient_name},\n\n"
                    f"We are pleased to inform you that Ticket No. {incident_num}, {prod_num} has been closed and root cause is {root_cause}.\n\n"
                    f"New OSC Jira ticket has been created for next action and please refer to the below Jira link(s):\n\n"
                    f"{osc_links}\n"
                    f"For details of the closed incident, please refer to the below Jira link:\n\n"
                    f"https://jira.devops.nonprod.empf.local/jira/browse/{prod_num}\n\n"
                    f"Thank you.\n\n"
                    f"Best regards,\n"
                    f"Liu Jun-Bing"
                )
            else:
                mail.Body = (
                    f"Dear {recipient_name},\n\n"
                    f"We are pleased to inform you that Ticket No.{incident_num}, {prod_num} has been resolved and root cause is {root_cause}.\n\n"
                    f"For detail, please refer to the below Jira link.\n\n"
                    f"https://jira.devops.nonprod.empf.local/jira/browse/{prod_num}\n\n"
                    f"Appreciate you could confirm the ticket is closed within the next 2 days.\n\n"
                    f"Best regards,\n"
                    f"Liu Jun-Bing"
                )
            try:
                # mail.Send()
                mail.Display()

            except Exception as e:
                print(f"{prod_num} - send email: {e}")
                logging.warning(f"{prod_num} - send email: {e}")

            # mail.Display()
            # time.sleep(60)  # 每封邮件间隔 60 秒


    elif channel == "Call Center" or channel == "Service Center":
        fixed_to = ()
        fixed_cc = ()
        # 构造主题
        email_subject = f"Ticket Resolved – {incident_num}, {prod_num}"
        # 固定收件人和抄送人
        if channel == "Call Center":
            fixed_to = (
                "Perry.CW.Kwok@pccw.com; Nesby.XZ.Li@pccw.com; Agnes.WM.Cheung@pccw.com; "
                "May.WM.Fa@pccw.com; Yuki.YL.Fan@pccw.com; Shirley.CY.Liang@pccw.com; "
                "Jeff.WB.Luo@pccw.com; Maggie.J.Wang@pccw.com; Stella.WS.Chen@pccw.com; "
                "grace.mn.liu@pccw.com; Kevin.CW.Wong@pccw.com; Vikey.HQ.Liu@pccw.com; "
                "Henry.CF.Tse@pccw.com; Sing.CS.Pang2@pccw.com"
            )
        elif channel == "Service Center":
            fixed_to = (
                "<Ryan.MY.Lai@pccw.com>; <Derek.KM.Siu@hkcsl.com>;"
            )

        if channel == "Call Center":
            fixed_cc = (
                "Stanislaus.SC.Tsang@pccw.com; Leo.CL.Chan@pccw.com; Ryan.MY.Lai@pccw.com; "
                "Yiki.YK.Kwan@pccw.com; Brian.CM.Kwok@pccw.com; im@support.empf.org.hk; "
                "Michael.L.Fan@pccw.com; May.XM.He@pccw.com; Tobia.Feng@pccw.com"
            )
        elif channel == "Service Center":
            fixed_cc = (
                "<Stanislaus.SC.Tsang@pccw.com>; <Leo.CL.Chan@pccw.com>; <Ryan.MY.Lai@pccw.com>;  "
                "<Derek.KM.Siu@hkcsl.com>;  <Yiki.YK.Kwan@pccw.com>;  <Brian.CM.Kwok@pccw.com>; "
                "<Tobia.Feng@pccw.com> ; <im@support.empf.org.hk>"
            )

        # 创建邮件
        mail = outlook.CreateItem(0)
        mail.To = fixed_to
        mail.CC = fixed_cc
        mail.Subject = email_subject
        name = ''
        if channel == "Call Center" :
            name = "Call Center"
        elif channel == "Service Center":
            name = "Service Center"


        # 判断 JiraStatus
        if jira_status == "INSUFFICIENT INFO":
            mail.HTMLBody = f"""
            <html>
            <body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
                <p>Dear {name},</p>

                <p>We acknowledge receipt of the captioned incident.<br>
                However, the information provided concerning this incident is insufficient for us to fully address the issue.
                To enable us to assist you further, appreciate you could provide additional details about the incident.</p>

                <p>此处有图片</p>

                <p>For detail, please refer to the below Jira link:<br>
                <a href="https://jira.devops.nonprod.empf.local/jira/browse/{prod_num}">
            ps://jira.devops.nonprod.empf.local/jira/browse/{prod_num}
                </a></p>

                <p>Once we receive the necessary information, we will be able to look into this matter and provide an appropriate response.</p>

                <p>If we do not receive any feedback from you regarding this additional information within the next 2 days, we will consider the ticket as closed.</p>

                <p>Best regards,<br>
                Jun-Bing, Liu</p>
            </body>
            </html>
            """
            embed_element_screenshot_in_email(driver, mail)
            mail.Display()

        elif jira_status == "Resolved" or jira_status == "Rejected" or jira_status == "Proposed To Close" or jira_status == "RELEASED":

            url = base_url_prefix = "https://jira.devops.nonprod.empf.local/jira/browse/" + prod_num
            driver.get(url)

            osc_ticket: list[WebElement] = driver.find_elements(By.CSS_SELECTOR,".links-list .link-content a[data-issue-key^='OSC-']")

            if osc_ticket:
                if len(osc_ticket) > 5:
                    element = driver.find_element(By.ID, "show-more-links-link")
                    # 滚动到该元素
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                    time.sleep(0.5)
                    # 点击元素
                    ActionChains(driver).move_to_element(element).click().perform()

                osc_links = ""
                for osc in osc_ticket:
                    osc_links += f"https://jira.devops.nonprod.empf.local/jira/browse/{osc.text}\n"

                mail.Subject = f"Ticket Closed – {incident_num}, {prod_num}"

                mail.Body = (
                    f"Dear {name},\n\n"
                    f"We are pleased to inform you that Ticket No. {incident_num}, {prod_num} has been closed and root cause is {root_cause}.\n\n"
                    f"New OSC Jira ticket has been created for next action and please refer to the below Jira link(s):\n\n"
                    f"{osc_links}\n"
                    f"For details of the closed incident, please refer to the below Jira link:\n\n"
                    f"https://jira.devops.nonprod.empf.local/jira/browse/{prod_num}\n\n"
                    f"Thank you.\n\n"
                    f"Best regards,\n"
                    f"Liu Jun-Bing"
                )
                try:
                    # mail.Send()
                    mail.Display()
                except Exception as e:
                    print(f"{prod_num} - send email: {e}")
                    logging.warning(f"{prod_num} - send email: {e}")
                # mail.Display()
            else:
                mail.HTMLBody = f"""
                <html>
                <body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
                    <p>Dear {name},</p>

                    <p>We are pleased to inform you that Ticket No. <strong>{incident_num}</strong>, <strong>{prod_num}</strong> has been resolved and root cause is <strong>{root_cause}</strong>.</p>

                    <p>Please recheck, thanks.</p>

                    <p><strong>此处有图片</strong></p>

                    <p>Appreciate you could confirm the ticket is closed within the next 2 days.</p>

                    <p>Best regards,<br>
                    Liu Jun-Bing</p>
                </body>
                </html>
                """
                embed_element_screenshot_in_email(driver, mail)
                mail.Display()


