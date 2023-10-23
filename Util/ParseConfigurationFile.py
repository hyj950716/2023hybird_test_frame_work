# -*- coding: utf-8 -*-
"""
@Time ： 2023/5/14 12:23
@Auth ： 胡英俊(请叫我英俊)
@File ：ParseConfigurationFile.py
@IDE ：PyCharm
@Motto：Never stop learning
"""
from configparser import ConfigParser
from Config.ProjVar import page_element_locator_file_path

class ParseConfigFile():

    def __init__(self):
        self.cf = ConfigParser()
        self.cf.read(page_element_locator_file_path,encoding="utf-8")

    def get_section_options(self,section_name):
        try:
            #把所有的选项全部放到字典中
            options_dict = dict(self.cf.items(section_name))
        except Exception:
            print("读取配置文件section:%s 出现异常！" %section_name)
            return None
        return options_dict

    def get_option_value(self,section_name,option_name):
        try:
            value = self.cf.get(section_name,option_name)
        except Exception:
            print("读取配置文件section:%s的选项%s 出现异常！" %(section_name,option_name))
            return None
        return value

if __name__ == "__main__":
    pc = ParseConfigFile()
    print(pc.get_section_options("login"))
    print(pc.get_option_value("contact","name_input_box"))
    print(pc.get_section_options("login1"))
    print(pc.get_option_value("contact1","name_input_box"))
