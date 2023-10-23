from Util.Excel import Excel
from Config.ProjVar import *
from KeyWord.KeyWord import *
from Util.Log import  *
import traceback
from Util.TakePic import *
from Util.TimeUtil import *
from Util.ParseConfigurationFile import *
import re

#从测试数据sheet中读取所有的测试数据
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
    return test_data

#根据测试步骤sheet，执行关键字的测试步骤
def execute_test_steps(test_step_sheet_name,test_result_sheet_name,head_line_flag=True):
    test_result = "成功"
    pc = ParseConfigFile()
    global wb
    wb.set_sheet_by_name(test_step_sheet_name)
    test_step_data = wb.get_all_rows_values()
    wb.set_sheet_by_name(test_result_sheet_name)
    if head_line_flag:
        wb.write_line(test_step_data[0], background_color="008000")  # 写一个表头
    for row_no in range(1, len(test_step_data)):
        key_word = test_step_data[row_no][keyword_col_no]
        locator_xpath_exp = test_step_data[row_no][locator_xpath_exp_col_no]
        if isinstance(locator_xpath_exp,str)  and  "||" in locator_xpath_exp :
            section_name = locator_xpath_exp.split("||")[0]
            element_name = locator_xpath_exp.split("||")[1]
            try:
                locator_xpath_exp = pc.get_option_value(section_name,element_name)
            except Exception as e:
                info("%s %s %s" %(section_name,element_name,"没有读取到对应的定位表达式"))
        value = test_step_data[row_no][value_col_no]
        if "$define" in key_word:
            test_step_result = execute_test_steps(locator_xpath_exp, value,head_line_flag=False)
            test_step_data[row_no][test_step_execute_result_col_no] = test_step_result
            continue
        # print(key_word,locator_xpath_exp,value)
        print("--------------------------------------------")

        # 情况1：func()
        if locator_xpath_exp is None and value is None:
            command = key_word + "()"
        # 情况2：func(arg1)
        elif locator_xpath_exp is None or value is None:
            if locator_xpath_exp is not None:
                command = key_word + '("%s")' % locator_xpath_exp
            else:
                command = key_word + '("%s")' % value
        # 情况3：func(arg1,arg2)
        else:
            command = key_word + '("%s","%s")' % (locator_xpath_exp, value)

        print(command)
        try:
            test_step_data[row_no][executed_time_col_no] = get_date_time()
            if "open_browser" in command:
                driver = eval(command)
            else:
                eval(command)
            test_step_data[row_no][test_step_execute_result_col_no] = "成功"
        except Exception as e:
            test_result = "失败"
            test_step_data[row_no][test_step_execute_result_col_no] = "失败"
            test_step_data[row_no][exception_info_col_no] = traceback.format_exc()
            pic_path = take_screenshot(driver)
            test_step_data[row_no][exception_screen_shot_path_col_no] = pic_path
            info("测试步骤：" + command)
            info("异常信息：" + traceback.format_exc())
        wb.write_line(test_step_data[row_no])
    wb.save()
    return test_result

#使用某一行测试数据，来执行关键字的测试步骤
def execute_test_steps_by_a_test_data_dict(test_step_sheet_name,test_result_sheet_name,test_data_dict,head_line_flag=True):
    test_result = "成功"
    pc = ParseConfigFile()
    global wb
    wb.set_sheet_by_name(test_step_sheet_name)
    test_step_data = wb.get_all_rows_values()
    wb.set_sheet_by_name(test_result_sheet_name)
    if head_line_flag:
        wb.write_line(test_step_data[0], background_color="008000")  # 写一个表头
    for row_no in range(1, len(test_step_data)):
        key_word = test_step_data[row_no][keyword_col_no]
        locator_xpath_exp = test_step_data[row_no][locator_xpath_exp_col_no]
        if isinstance(locator_xpath_exp,str)  and  "||" in locator_xpath_exp :
            section_name = locator_xpath_exp.split("||")[0]
            element_name = locator_xpath_exp.split("||")[1]
            try:
                locator_xpath_exp = pc.get_option_value(section_name,element_name)
            except Exception as e:
                info("%s %s %s" %(section_name,element_name,"没有读取到对应的定位表达式"))
        value = test_step_data[row_no][value_col_no]
        if isinstance(value,str) and re.search(r"\$\{.*?\}", value):
            var_name = re.search(r"\$\{(.*?)\}", value).group(1)
            if var_name in test_data_dict.keys():
                value = test_data_dict[var_name]
            else:
                info(var_name+"在字典test_data_dict:%s 中不存在" %test_data_dict)
        if "$define" in key_word:
            test_step_result = execute_test_steps(locator_xpath_exp, value,head_line_flag=False)
            test_step_data[row_no][test_step_execute_result_col_no] = test_step_result
            continue
        # print(key_word,locator_xpath_exp,value)
        print("--------------------------------------------")

        # 情况1：func()
        if locator_xpath_exp is None and value is None:
            command = key_word + "()"
        # 情况2：func(arg1)
        elif locator_xpath_exp is None or value is None:
            if locator_xpath_exp is not None:
                command = key_word + '("%s")' % locator_xpath_exp
            else:
                command = key_word + '("%s")' % value
        # 情况3：func(arg1,arg2)
        else:
            command = key_word + '("%s","%s")' % (locator_xpath_exp, value)

        print(command)
        try:
            test_step_data[row_no][executed_time_col_no] = get_date_time()
            if "open_browser" in command:
                driver = eval(command)
            else:
                eval(command)
            test_step_data[row_no][test_step_execute_result_col_no] = "成功"
        except Exception as e:
            test_result = "失败"
            test_step_data[row_no][test_step_execute_result_col_no] = "失败"
            test_step_data[row_no][exception_info_col_no] = traceback.format_exc()
            pic_path = take_screenshot(driver)
            test_step_data[row_no][exception_screen_shot_path_col_no] = pic_path
            info("测试步骤：" + command)
            info("异常信息：" + traceback.format_exc())
        wb.write_line(test_step_data[row_no])
    wb.save()
    return test_result

def execute_keyword_test(test_step_sheet_name,test_result_sheet_name):
    test_result = execute_test_steps(test_step_sheet_name, test_result_sheet_name)
    return test_result

def execute_hybrid_test(test_step_sheet_name,test_data_sheet_name,test_result_sheet_name):
    test_result = "成功"
    test_data_list = get_test_data(test_data_sheet_name)
    for test_data_dict in test_data_list:
        if test_data_dict["是否执行"]=="y" or test_data_dict["是否执行"]=="Y":
            test_data_dict["执行时间"]=get_date_time()
            test_data_dict["测试结果"] = execute_test_steps_by_a_test_data_dict(test_step_sheet_name,test_result_sheet_name,test_data_dict)
            table_head = [key for key in test_data_dict.keys()]
            line = [value for value in test_data_dict.values()]
            wb.set_sheet_by_name(test_result_sheet_name)
            wb.write_line(table_head,background_color="018000")
            wb.write_line(line)
            if "失败" in  test_data_dict["测试结果"]:
                test_result = "失败"
    return test_result

#execute_hybrid_test("登录1","登录测试数据","测试结果")

if __name__ =="__main__":
    wb = Excel(test_data_file_path )
    wb.set_sheet_by_name("测试用例")
    test_cases = wb.get_all_rows_values()
    for row_no in range(1, len(test_cases)):
        # 读出测试用例是否执行的标志位
        test_case_if_executed_flag = test_cases[row_no][test_case_if_executed_flag_col_no]
        if "y" not in test_case_if_executed_flag.lower():
            continue
        # 读出测试步骤的所在sheet名称
        test_step_sheet_name = test_cases[row_no][test_step_sheet_name_col_no]
        # 读出测试结果的所在sheet名称
        test_result_sheet_name = test_cases[row_no][test_result_sheet_name_col_no]
        #测试数据所在的sheet名称
        test_data_sheet_name = test_cases[row_no][test_data_sheet_name_col_no]
        #获取当前时间，写入到当前测试用例行中的测试时间单元格
        test_cases[row_no][test_executed_time_col_no]=get_date_time()
        # 要按照关键字框架来执行
        if "n"== test_data_sheet_name:
            test_result = execute_keyword_test(test_step_sheet_name ,test_result_sheet_name)
        #按照混合框架模式来执行
        else:
            test_result =execute_hybrid_test(test_step_sheet_name,test_data_sheet_name,test_result_sheet_name)
        #写入到当前测试用例行中的测试结果单元格
        test_cases[row_no][test_result_col_no]=test_result
        #设定要操作的sheet名称
        wb.set_sheet_by_name("测试结果")
        #写入测试用例sheet的表头
        wb.write_line(test_cases[0],background_color="018000")
        #写入当前测试用例行的所有内容到测试结果sheet中
        wb.write_line(test_cases[row_no])



