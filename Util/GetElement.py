# -*- coding: utf-8 -*-
"""
@Time ： 2023/5/14 12:14
@Auth ： 胡英俊(请叫我英俊)
@File ：GetElement.py
@IDE ：PyCharm
@Motto：Never stop learning
"""
# encoding=utf-8
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver


# 获取单个页面元素
def get_element(driver, locator_exp):
    try:
        element = WebDriverWait(driver, 30).until(
            lambda x: x.find_element(by="xpath", value=locator_exp))
        return element
    except Exception as err:
        raise err


# 获取多个页面元素
def get_elements(driver, locator_exp):
    try:
        elements = WebDriverWait(driver, 30).until(
            lambda x: x.find_elements(by="xpath", value=locator_exp))
        return elements
    except Exception as err:
        raise err


if __name__ == "__main__":
    driver = webdriver.Chrome(executable_path="e:\\chromedriver.exe")
    driver.get("http://mail.126.com")
    iframe = get_element(driver, '//div[@id="loginDiv"]/iframe')
    print(iframe)
    links = get_elements(driver, '//a[@href]')
    print(len(links))
