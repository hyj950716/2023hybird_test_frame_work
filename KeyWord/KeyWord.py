# -*- coding: utf-8 -*-
"""
@Time ： 2023/5/21 12:06
@Auth ： 胡英俊(请叫我英俊)
@File ：KeyWord.py
@IDE ：PyCharm
@Motto：Never stop learning
"""
from selenium import webdriver
from Config.ProjVar import *
import os
import time
from selenium.webdriver.common.by import By

driver = ""


def open_browser(browser_name):
    global driver
    if "chrome" in browser_name.lower():
        options = webdriver.ChromeOptions()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        driver = webdriver.Chrome(executable_path=chrome_driver_path, chrome_options=options)
    elif "ie" in browser_name.lower():
        driver = webdriver.Ie(executable_path=ie_driver_path)
    else:
        driver = webdriver.Firefox(executable_path=firefox_driver_path)
    return driver


def visit(url):
    global driver
    driver.get(url)


def sleep(seconds):
    seconds = int(seconds)
    time.sleep(seconds)


def switch_in_frame(frame_xpath_exp):
    global driver
    frame = driver.find_element(By.XPATH, frame_xpath_exp)
    driver.switch_to.frame(frame)


def switch_out_frame():
    global driver
    driver.switch_to.default_content()


def input(element_xpath_exp, content):
    global driver
    input_box = driver.find_element(By.XPATH, element_xpath_exp)
    input_box.clear()
    input_box.send_keys(content)


def click(element_xpath_exp):
    global driver
    element = driver.find_element(By.XPATH, element_xpath_exp)
    element.click()


def assert_word(word):
    global driver
    assert word in driver.page_source


def quit():
    global driver
    driver.quit()


if __name__ == "__main__":
    driver = open_browser("chrome")
    visit("https://www.126.com")
    switch_in_frame("//div[@id='loginDiv']/iframe")
    input("//input[@tabindex='1']", "hxf950716")
    input("//*[@tabindex=2]", "Hxf950716!")
    click("//a[@id='dologin']")
    sleep(7)
    click("//span[.='写 信']")
    time.sleep(2)
    input("//input[@aria-label='收件人地址输入框，请输入邮件地址，多人时地址请以分号隔开']", "testman1980@126.com")
    input("//input[contains(@id,'subjectInput')]", "今天天气不错")
    sleep(2)
    input("//input[@type='file']", "e:\\a.py")
    switch_in_frame("//iframe[@tabindex='1']")
    input("//body/p", "好好学习天天向上！")
    switch_out_frame()
    click("//span[.='发送']")
    sleep(2)
    assert_word("邮件发送成功")
    click("//span[@id='spnUid']")
    click("//div[.='退出登录']")
    quit()
