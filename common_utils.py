from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.common.by import By
from dotenv import load_dotenv
import os
import time
import win32com.client

def get_excel_file_url():
    # 加载 .env 文件
    load_dotenv()

    # 读取变量
    excel_file = os.getenv("EXCEL_FILE")
    return excel_file

def embed_element_screenshot_in_email(driver, mail, element_id="issue_actions_container"):
    """
    截图指定元素并嵌入到Outlook邮件正文中（HTML + CID）。

    参数:
    driver: Selenium WebDriver 实例
    mail: Outlook MailItem 对象
    element_id: 要截图的元素ID，默认 'issue_actions_container'
    """
    # 截图保存路径
    screenshot_path = os.path.join(os.getcwd(), f"{element_id}.png")

    # 查找元素并截图
    element = driver.find_element(By.ID, element_id)

    # 使用 JavaScript 滚动元素到可视区域
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)

    element.screenshot(screenshot_path)

    # 添加图片到邮件并设置CID
    attachment = mail.Attachments.Add(screenshot_path)
    attachment.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "issue_img"
    )

    # 在HTML正文中插入图片
    mail.HTMLBody = mail.HTMLBody.replace("此处有图片", '<img src="cid:issue_img" style="max-width:600px;">')

def check_subject_in_sent(subject_keyword, max_check=500):
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    outlook = outlook_app.GetNamespace("MAPI")
    sent_folder = outlook.GetDefaultFolder(5)  # 5 = Sent Items
    messages = sent_folder.Items
    messages.Sort("[SentOn]", True)  # 按时间降序排序

    checked = 0
    for msg in messages:
        if checked >= max_check:
            break
        checked += 1
        try:
            if msg.Subject and subject_keyword.lower() in msg.Subject.lower():
                # print(f"✅ 已发送，匹配主题: {msg.Subject}")
                return True
        except Exception as e:
            continue  # 忽略无法读取的邮件项

    return False

def wait_for_element(driver, by, value, condition="visible", timeout=10):
    """
    显式等待某个元素满足指定条件。

    参数:
        driver: Selenium WebDriver 实例
        by: 定位方式，如 By.ID, By.XPATH, By.CSS_SELECTOR 等
        value: 定位值
        condition: 等待条件，可选值：
            - "visible": 元素可见
            - "clickable": 元素可点击
            - "present": 元素存在于 DOM 中
        timeout: 最长等待时间（秒）

    返回:
        WebElement 对象（如果找到），否则抛出 TimeoutException
    """
    wait = WebDriverWait(driver, timeout)

    if condition == "visible":
        return wait.until(EC.visibility_of_element_located((by, value)))
    elif condition == "clickable":
        return wait.until(EC.element_to_be_clickable((by, value)))
    elif condition == "present":
        return wait.until(EC.presence_of_element_located((by, value)))
    else:
        raise ValueError(f"未知的等待条件: {condition}")


def login_jira():
    # 加载 .env 文件
    load_dotenv()

    # 读取变量
    jira_username = os.getenv("JIRA_USERNAME")
    jira_password = os.getenv("JIRA_PASSWORD")
    driver = webdriver.Chrome()
    driver.get("https://jira.devops.nonprod.empf.local/jira/login.jsp")
    driver.maximize_window()
    driver.implicitly_wait(10)

    driver.find_element(By.ID, "login-form-username").send_keys(jira_username)
    driver.find_element(By.ID, "login-form-password").send_keys(jira_password)
    driver.find_element(By.ID, "login-form-submit").click()
    time.sleep(1)
    return driver

def confirm_box(driver):
    # 注入自定义弹窗
    js_code = """
    var modal = document.createElement('div');
    modal.id = 'customModal';
    modal.style.position = 'fixed';
    modal.style.top = '50%';
    modal.style.left = '50%';
    modal.style.transform = 'translate(-50%, -50%)';
    modal.style.backgroundColor = 'white';
    modal.style.padding = '20px';
    modal.style.border = '2px solid black';
    modal.style.zIndex = '9999';

    var message = document.createElement('p');
    message.innerText = 'Do you want to continue? Please enter a remark:';
    modal.appendChild(message);

    var input = document.createElement('input');
    input.type = 'text';
    input.id = 'remarkInput';
    modal.appendChild(input);

    var yesBtn = document.createElement('button');
    yesBtn.innerText = 'Yes';
    yesBtn.onclick = function() {
        window.userResponse = { choice: 'Yes', remark: document.getElementById('remarkInput').value };
        document.body.removeChild(modal);
    };
    modal.appendChild(yesBtn);

    var noBtn = document.createElement('button');
    noBtn.innerText = 'No';
    noBtn.style.marginLeft = '10px';
    noBtn.onclick = function() {
        window.userResponse = { choice: 'No', remark: document.getElementById('remarkInput').value };
        document.body.removeChild(modal);
    };
    modal.appendChild(noBtn);

    document.body.appendChild(modal);
    """

    driver.execute_script(js_code)

    # 等待用户操作（轮询直到用户点击按钮）
    user_response = None
    while user_response is None:
        try:
            user_response = driver.execute_script("return window.userResponse;")
        except:
            pass
        # time.sleep(1)

    print("User clicked:", user_response['choice'])
    print("User remark:", user_response['remark'])

    return {
        "choice": user_response['choice'],
        "remark": user_response['remark'],
    }