# -*- coding: utf-8 -*-
"""
@Time ： 2023/5/14 11:41
@Auth ： 胡英俊(请叫我英俊)
@File ：Log.py
@IDE ：PyCharm
@Motto：Never stop learning
"""
"""封装的操作日志函数"""

import logging
import logging.config

from Config.ProjVar import log_config_path

#使用日志的配置文件来生成日志对象
logging.config.fileConfig(log_config_path)

#选择日志打印方式
logger = logging.getLogger("example01")

def debug(message):
    logger.debug(message)

def info(message):
    logger.info(message)

def error(message):
    logger.error(message)

if __name__ == "__main__":
    debug("今天天气不错")
    info("信息流")
    error("有个错误！")

