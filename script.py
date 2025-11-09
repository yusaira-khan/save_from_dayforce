from selenium import webdriver
from selenium.webdriver.common.by import By
import subprocess
import os
from selenium.webdriver.common.print_page_options import PrintOptions
import base64
import time
import datetime
import re
username, password, company = [subprocess.Popen(f"op read \"op://Private/Dayforce/{aa}\"",stdout=subprocess.PIPE).stdout.read().decode("utf-8") for aa in ["username", "password", "company"]]
p = PrintOptions()
p.page_ranges = ["2"]
p.shrink_to_fit=False
p.margin_top=0
p.margin_left=0
p.margin_bottom=0
p.margin_right=0

driver = webdriver.Firefox()
driver.get("https://www.dayforcehcm.com")
driver.find_element(By.ID,"txtCompanyId").send_keys(company)
driver.find_element(By.ID,"MainContent_loginUI_cmdLogin").click()
driver.find_element(By.ID,"txtNewUserName").send_keys(username)
driver.find_element(By.ID,"MainContent_loginUI_cmdLogin").click()
driver.find_element(By.ID,"txtNewUserPass").send_keys(password)
driver.find_element(By.ID,"MainContent_loginUI_cmdLogin").click()

time.sleep(8)
driver.find_element(By.ID,"FeatureDetail_0").click()
time.sleep(3)

from_txt=driver.find_element(By.ID,"dateBoxFrom")
from_txt.clear()
from_txt.send_keys(os.environ["from"])

to_txt=driver.find_element(By.ID,"dateBoxTo")
to_txt.clear()
to_txt.send_keys(os.environ["to"])

driver.find_element(By.CLASS_NAME,"dijitButton").click()
time.sleep(2)
driver.find_element(By.CLASS_NAME,"linkButtonClass").click()
time.sleep(3)

def get_statement_idx():
    return int(re.match("\\d+", driver.find_element(By.ID, "Earning_Dialog_Header_Top").text).group())

def click_previous():
    driver.find_element(By.CLASS_NAME, "statementNavButton").click()

def save_page():
    date_header = driver.find_element(By.ID, "Earning_Dialog_Header").text
    date_txt = (datetime.datetime.strptime(
        re.search("\\d+/\\d+/\\d+", date_header).group(), "%m/%d/%Y")
    .strftime( "%Y-%m-%d"))
    add_ext = "" if re.search("Additional", date_header) is None else ".extra"
    with open(f"{date_txt}{add_ext}.pdf", "wb") as f: f.write(base64.b64decode(driver.print_page(p)))

while True:
    idx = get_statement_idx()
    print(f"statement {idx}")
    save_page()
    if idx == 1:
        break
    else:
        click_previous()
        time.sleep(5)


# input("Press Enter to continue...")
driver.close()