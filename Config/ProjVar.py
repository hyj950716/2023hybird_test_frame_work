# -*- coding: utf-8 -*-
"""
@Time ： 2023/5/15 14:30
@Auth ： 胡英俊(请叫我英俊)
@File ：ProjVar.py
@IDE ：PyCharm
@Motto：Never stop learning
"""
import os

# print(__file__)#获得文件当前的绝对路径
# print(os.path.dirname(__file__))

# 获取当前工程所在的目录
proj_path = os.path.dirname(os.path.dirname(__file__))

# 获取日志配置文件的所在路径
log_config_path = os.path.join(proj_path, "Config\\Logger.conf")

# 页面元素定位表达式配置文件的所在路径
page_element_locator_file_path = os.path.join(proj_path, "Config\\PageElementLocator.ini")

# 获取excel数据文件的路径
test_data_file_path = os.path.join(proj_path, "Data\\测试数据文件.xlsx")

# 各个浏览器的驱动位置，可以配置成自己常用的浏览器驱动位置，主要驱动版本要和浏览器版本保持一致
chrome_driver_path = "e:\\chromedriver.exe"
ie_driver_path = "e:\\IEDriverServer.exe"
firefox_driver_path = "e:\\geckodriver.exe"

keyword_col_no = 2
locator_xpath_exp_col_no = 3
value_col_no = 4
executed_time_col_no = 5
test_step_execute_result_col_no = 6
exception_info_col_no = 7
exception_screen_shot_path_col_no = 8

test_step_sheet_name_col_no = 3
test_result_sheet_name_col_no = 4
test_executed_time_col_no = 5
test_result_col_no = 6
test_case_if_executed_flag_col_no = 7
test_data_sheet_name_col_no = 8

if __name__ == "__main__":
    print(proj_path)
    print(log_config_path)
    print(page_element_locator_file_path)
    print(test_data_file_path)
