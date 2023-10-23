# -*- coding: utf-8 -*-
"""
@Time ： 2023/5/21 9:48
@Auth ： 胡英俊(请叫我英俊)
@File ：DirUtil.py
@IDE ：PyCharm
@Motto：Never stop learning
"""
import os
from Util.TimeUtil import *
from Config.ProjVar import proj_path

def make_date_dir():
    year = get_year()
    year_dir_path = os.path.join(proj_path,"ScreenShot",year)
    if not os.path.exists(year_dir_path):
        os.mkdir(year_dir_path)

    month = get_month()
    month_dir_path = os.path.join(year_dir_path,month)
    if not os.path.exists(month_dir_path):
        os.mkdir(month_dir_path)

    day = get_day()
    day_dir_path = os.path.join(month_dir_path,day)
    if not os.path.exists(day_dir_path):
        os.mkdir(day_dir_path)
    return day_dir_path

if __name__ == "__main__":
    print(make_date_dir())
