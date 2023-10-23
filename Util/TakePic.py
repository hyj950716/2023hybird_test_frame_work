# -*- coding: utf-8 -*-
"""
@Time ： 2023/5/21 9:47
@Auth ： 胡英俊(请叫我英俊)
@File ：TakePic.py
@IDE ：PyCharm
@Motto：Never stop learning
"""
from Util.DirUtil import  make_date_dir
from Util.TimeUtil import current_time
from selenium import webdriver
import os
def take_screenshot(driver):
    dir_path = make_date_dir()#拿到截屏的保存目录
    file_name = current_time()+".png"
    file_path = os.path.join(dir_path,file_name)
    driver.get_screenshot_as_file(file_path)
    return file_path

if __name__ =="__main__":
    driver = webdriver.Chrome(executable_path="e:\\chromedriver.exe")
    driver.get("https://www.baidu.com")
    print(take_screenshot(driver))
    driver.close()
