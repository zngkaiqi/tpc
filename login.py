# 人事行政系統自動化
# 功能包含: 自動化登入、代理假單、查詢刷卡資料、假單登記

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoSuchWindowException
from selenium.common.exceptions import NoAlertPresentException
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.ie.service import Service
from datetime import datetime, timedelta
from time import sleep
import subprocess

offtime = None
offdt_date_path = r"C:\tpc\offdt_date.txt"
offdt_time_path = r"C:\tpc\offdt_time.txt"
offregistered = False

# BgInfo: 在桌面顯示下班時間，須另外下載
bginfo_path = r"C:\Program Files\BGInfo\Bginfo64.exe"
bginfoconfig_path = r"C:\Program Files\BGInfo\config.bgi"

def accept_all_alert():
    while True:
        try:
            driver.switch_to.alert.accept()
        except NoAlertPresentException:
            break

# Initialize
service = Service('./IEDriverServer.exe')
options = webdriver.IeOptions()
options.attach_to_edge_chrome = True
# options.initial_browser_url = "http://sso.taipower.com.tw/wps/portal/0/login/!ut/p/z1/04_Sj9CPykssy0xPLMnMz0vMAfIjo8zi_QwNLTxMLAz83Q0tDQwCLUwCnVwN_YwNAk31wwkpiAJKG-AAjgb6BbmhigBUllSD/dz/d5/L2dBISEvZ0FBIS9nQSEh/"
options.initial_browser_url = "http://sso.taipower.com.tw/"
driver = webdriver.Ie(service=service, options=options)
wait = WebDriverWait(driver, 5)

# SSO: 使用記住的帳號密碼
for _ in range(5):
    if 'sso' not in driver.current_url:
        driver.get("http://sso.taipower.com.tw/")
    userid = wait.until(EC.element_to_be_clickable((By.NAME, "userid")))
    # userid.send_keys("076531")
    userid.send_keys(Keys.ARROW_DOWN)
    userid.send_keys(Keys.ARROW_DOWN)
    userid.send_keys(Keys.ENTER)
    driver.execute_script("goSubmit();")
    wait.until(EC.url_contains("sso.taipower.com.tw/wps/myportal"))
    if 'myportal' in driver.current_url: break
    else: driver.refresh()
        
# 人事行政
driver.get("http://sso.taipower.com.tw/wps/myportal/!ut/p/z1/04_Sj9CPykssy0xPLMnMz0vMAfIjo8zi_QwNLTxMLAz83Q0tDQwCLUwCnVwN_YwNQs30wwkpiAJKG-AAjgb6BbmhigDgayq3/dz/d5/L2dJQSEvUUt3QS80TmxFL1o2X04xMThINDgwT0cxOTAwUTg0UUJFMU4zR0kx/")
accept_all_alert()
for handle in driver.window_handles:
    driver.switch_to.window(handle)
    accept_all_alert()
    if 'sso' in driver.current_url:
        driver.close()
driver.switch_to.window(driver.window_handles[0])
wait.until(EC.title_contains('人事行政'))

# 刷卡資料更新
def offdt_update():
    global offtime
    if offtime is not None:
        if offregistered:            
            with open(offdt_date_path, "w") as f:
                f.write(f"({datetime.now().month}/{datetime.now().day})")
        return
    driver.switch_to.new_window('tab')
    driver.get("http://stpc02402199.taipower.com.tw/PBS/EI/EI0100MainClassX.asp")
    driver.switch_to.frame("EItop")
    driver.execute_script("submitcheck6();") # 刷卡資料查詢.click()
    driver.switch_to.default_content()
    driver.switch_to.frame("bottom")
    try:
        offtime = datetime.strptime(driver.find_element(By.ID, "dgCIN__ctl2_Label2").text, "%H:%M") + timedelta(hours=8, minutes=40)
        offtime = datetime.now().replace(hour=offtime.hour, minute=offtime.minute)
        with open(offdt_time_path, "w") as f:
            f.write(offtime.strftime('%H:%M'))
    except NoSuchElementException:
        with open(offdt_time_path, "w") as f:
            f.write('N/A')
    with open(offdt_date_path, "w") as f:
        f.write(f"{datetime.now().month}/{datetime.now().day}")
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    subprocess.run([bginfo_path, bginfoconfig_path, '/timer:0', '/silent', '/nolicprompt'])                    

# 代理人簽核
def agent_approve():
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        if '人事行政' in driver.title:
            break
    driver.switch_to.frame("bottom")
    try:
        elem = driver.find_element(By.PARTIAL_LINK_TEXT, "代理簽核")
        if int(''.join(filter(lambda x: x.isdigit(), elem.text))) > 0:
            driver.execute_script("arguments[0].click();", elem)
            for elem in driver.find_elements(By.XPATH, "//input[@value='ok']"):
                driver.execute_script("arguments[0].checked = true;", elem)
            driver.execute_script("document.forms[0].submit()")
            sleep(1)
    except NoSuchElementException:
        pass
    driver.switch_to.default_content()

# 請假申請
def off_reg():
    start_time = datetime.now() + timedelta(minutes=3)
    # if start_time > datetime.now().replace(hour=12, minute=0): return
    end_time = offtime if (not offtime is None) else start_time.replace(hour=17, minute=10)
    tdelta = end_time - start_time - timedelta(minutes=40) if start_time < datetime.now().replace(hour=12, minute=0) else end_time - start_time
    time_total = int(tdelta.total_seconds()/3600) if tdelta.total_seconds() % 3600 == 0 else int(tdelta.total_seconds()/3600) + 1
    driver.switch_to.new_window('tab')
    driver.get("http://stpc02402199.taipower.com.tw/PBS/EI/EI0100MainClassX.asp")
    driver.switch_to.frame("EItop")
    driver.execute_script("arguments[0].click();", driver.find_element(By.LINK_TEXT, "差假管理"))
    driver.switch_to.default_content()
    wait.until(EC.title_is("差假管理系統"))
    driver.switch_to.frame("bottom")
    driver.switch_to.frame("frmTools")
    driver.execute_script("arguments[0].click();", driver.find_element(By.LINK_TEXT, "申請"))
    driver.switch_to.parent_frame()
    driver.switch_to.frame("frmContent")
    driver.switch_to.frame("frmContent")
    driver.execute_script("arguments[0].value='22公出';", driver.find_element(By.XPATH, "//select[@name='OffClass']"))
    driver.execute_script("arguments[0].value='處本部';", driver.find_element(By.XPATH, "//input[@name='tmp_explace']"))
    driver.execute_script("arguments[0].value='034239沈煒傑';", driver.find_element(By.XPATH, "//select[@name='AgentList']"))
    driver.execute_script("arguments[0].value='{}';".format(start_time.hour), driver.find_element(By.XPATH, "//select[@name='StartHO']"))
    driver.execute_script("arguments[0].value='{}';".format(start_time.minute), driver.find_element(By.XPATH, "//select[@name='StartMN']"))
    driver.execute_script("arguments[0].value='{}';".format(end_time.hour), driver.find_element(By.XPATH, "//select[@name='EndHO']"))
    driver.execute_script("arguments[0].value='{}';".format(end_time.minute), driver.find_element(By.XPATH, "//select[@name='EndMN']"))
    driver.execute_script("arguments[0].value='{}';".format(0 if time_total < 8 else 1), driver.find_element(By.XPATH, "//input[@name='TOTDAY']"))
    driver.execute_script("arguments[0].value='{}';".format(time_total if time_total < 8 else time_total - 8), driver.find_element(By.XPATH, "//select[@name='TOTHOUR']"))
    driver.execute_script("arguments[0].value='公文傳送及資控業務辦理。';", driver.find_element(By.XPATH, "//textarea[@name='Reason']"))
    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//input[@name='btnsubmit']"))
    sleep(1)
    accept_all_alert()
    if '已有其他有效假單' in driver.page_source:
        global offregistered
        offregistered = True
    else:
        sleep(200)
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

while True:
    try:
        agent_approve()
        offdt_update()
        if datetime.now().weekday() == 2 and not offregistered:
            off_reg()
        # off_reg()
        sleep(5)
        driver.refresh()
    except UnexpectedAlertPresentException:
        accept_all_alert()
    except (KeyboardInterrupt, NoSuchWindowException):
        driver.quit()
        break