import shutil

import pandas as pd
from datetime import datetime
from common_utils import *


# 清除缓存
gen_py_path = os.path.join(os.environ["LOCALAPPDATA"], "Temp", "gen_py")
if os.path.exists(gen_py_path):
    shutil.rmtree(gen_py_path)

# 读取 Excel
excel_file = get_excel_file_url()
sheet_name = "Sheet1"
df = pd.read_excel(excel_file, sheet_name=sheet_name, engine="openpyxl", header=0)
df.columns = ["RootCause", "Email", "Name","JiraStatus","Channel","Id", "JiraKey", "SendEmail","EmailStatus","SentDate","Remark"]


df["SendEmail"] = df["SendEmail"].astype(str)
df["EmailStatus"] = df["EmailStatus"].astype(str)
df["SentDate"] = df["SentDate"].astype(str)

# 先筛选出需要发送邮件的行
filtered_df = df[(df["SendEmail"] != "no") & (df["EmailStatus"] != "sent")]




for index, row in filtered_df.iterrows():
    JiraKey = str(row["JiraKey"])
    send_email = str(row["SendEmail"])
    remark = str(row["Remark"])
    if send_email == 'no':
        print(f"不需要发送：{JiraKey},{remark}")
        df.at[index, "EmailStatus"] = ""
        df.at[index, "SentDate"] = ""
        continue
    if check_subject_in_sent(JiraKey):
        df.at[index, "SendEmail"] = "yes"
        df.at[index, "EmailStatus"] = "sent"
        df.at[index, "SentDate"] = datetime.today().strftime('%m/%d/%Y')
        # df.at[index, "Remark"] = ""
        print(f"✅ 已发送：{JiraKey}")
    else:
        df.at[index, "SendEmail"] = "yes"
        df.at[index, "EmailStatus"] = ""
        df.at[index, "SentDate"] = ""
        print(f"❌ 未发送：{JiraKey}")

# 保存 Excel
df.to_excel(excel_file, sheet_name=sheet_name, index=False, engine="openpyxl")