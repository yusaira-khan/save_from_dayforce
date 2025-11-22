from selenium import webdriver
from selenium.webdriver.common.by import By
import subprocess
import os
import time
from selenium.webdriver.common.print_page_options import PrintOptions
import base64
import datetime
import re


def open_page(credentials):
    driver.get("https://www.dayforcehcm.com")
    driver.find_element(By.ID, "txtCompanyId").send_keys(credentials["company"])
    driver.find_element(By.ID, "MainContent_loginUI_cmdLogin").click()
    driver.find_element(By.ID, "txtNewUserName").send_keys(credentials["username"])
    driver.find_element(By.ID, "MainContent_loginUI_cmdLogin").click()
    driver.find_element(By.ID, "txtNewUserPass").send_keys(credentials["password"])
    driver.find_element(By.ID, "MainContent_loginUI_cmdLogin").click()

    time.sleep(14)
    driver.find_element(By.CLASS_NAME, "FeatureDetail").click()
    time.sleep(3)

    from_txt = driver.find_element(By.ID, "dateBoxFrom")
    from_txt.clear()
    from_txt.send_keys(os.environ["from"])

    to_txt = driver.find_element(By.ID, "dateBoxTo")
    to_txt.clear()
    to_txt.send_keys(os.environ["to"])

    driver.find_element(By.CLASS_NAME, "dijitButton").click()
    time.sleep(2)
    driver.find_element(By.CLASS_NAME, "linkButtonClass").click()
    time.sleep(3)


def get_statement_idx():
    return int(re.match("\\d+", driver.find_element(By.ID, "Earning_Dialog_Header_Top").text).group())


def click_previous():
    driver.find_element(By.CLASS_NAME, "statementNavButton").click()


def page_name():
    date_header = driver.find_element(By.ID, "Earning_Dialog_Header").text
    date_txt = (datetime.datetime.strptime(
        re.search("\\d+/\\d+/\\d+", date_header).group(), "%m/%d/%Y")
                .strftime("%Y-%m-%d"))
    add_ext = "" if re.search("Additional", date_header) is None else ".extra"
    return f"{date_txt}{add_ext}"


_opt = None


def opt():
    global _opt
    if _opt is None:
        _opt = PrintOptions()
        _opt.page_ranges = ["2"]
        _opt.shrink_to_fit = False
        _opt.margin_top = 0
        _opt.margin_left = 0
        _opt.margin_bottom = 0
        _opt.margin_right = 0
    return _opt


def save_pdf():
    content = base64.b64decode(driver.print_page(opt()))
    with open(f"{page_name()}.pdf", "wb") as f:
        f.write(content)


def save_html():
    content = driver.find_element(By.CLASS_NAME, "earningDetailPanel").get_attribute("outerHTML")
    with open(f"{page_name()}.html", "w") as f:
        f.write(content)


def save_pages():
    while True:
        idx = get_statement_idx()
        print(f"statement {idx}")
        save_html()
        if idx == 1:
            break
        else:
            click_previous()
            time.sleep(5)


def get_credentials():
    return {
        field:
            subprocess.Popen(
                f"op read \"op://Private/Dayforce/{field}\"",
                stdout=subprocess.PIPE)
            .stdout
            .read()
            .decode("utf-8")
        for field in ["username", "password", "company"]
    }

driver = webdriver.Firefox()
open_page(get_credentials())
save_pages()
driver.close()
