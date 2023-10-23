# -*- coding: utf-8 -*-
"""
@Time ： 2023/5/15 14:30
@Auth ： 胡英俊(请叫我英俊)
@File ：TimeUtil.py
@IDE ：PyCharm
@Motto：Never stop learning
"""
"""
封装常用调用的时间格式方法
"""
import time

def date():
    date = str(time.localtime().tm_year) + "年" + str(time.localtime().tm_mon) + "月" + str(
        time.localtime().tm_mday) + "日"
    return date

def current_time():
    now_time = str(time.localtime().tm_hour) + "时" + str(time.localtime().tm_min) + "分" + str(
        time.localtime().tm_sec) + "秒"
    return now_time

def get_date_time():
    return date()+" "+current_time()

def get_year():
    return str(time.localtime().tm_year) + "年"

def get_month():
    return str(time.localtime().tm_mon) + "月"

def get_day():
    return str( time.localtime().tm_mday) + "日"


if __name__ =="__main__":
    print(get_date_time())
    print(date())
    print(get_year())
    print(get_month())
    print(get_day())
    print(current_time())