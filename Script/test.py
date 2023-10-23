# -*- coding: utf-8 -*-
"""
@Time ： 2023/5/28 10:36
@Auth ： 胡英俊(请叫我英俊)
@File ：test.py
@IDE ：PyCharm
@Motto：Never stop learning
"""
from Config.ProjVar import test_data_file_path
from Util.Excel import Excel

wb = Excel(test_data_file_path)
def get_test_data(test_data_sheet_name):
    wb.set_sheet_by_name(test_data_sheet_name)
    rows = wb.get_all_rows_values()#读取表格中所有的数据行

    test_data =[]
    #从第二行开始遍历所有的数据行
    for row in rows[1:]:
        d = {}
        for col_no in range(len(rows[0])):#遍历第一行所有的列号
            key = rows[0][col_no]#可以获得每一个key
            value = row[col_no]#读取当前行和key对应的单元格值
            d[key] = value
        test_data.append(d)
    print(test_data)

print(get_test_data("登录测试数据"))

d={"pass_word":"1234abc"}
value = "  ${pass_word}  "
import re
if re.search(r"\$\{.*?\}",value):
    print("匹配到了！")
    var_name = re.search(r"\$\{(.*?)\}", value).group(1)
    print(var_name)
    print("从字典中读到的值：" ,d[var_name])
else:
    print("没有匹配到")



