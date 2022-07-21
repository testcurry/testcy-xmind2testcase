#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import csv
import pandas as pd  # 使用pandas包来写入excel的，需要引入
from pandas.io.excel import ExcelWriter
import xlwt
import logging
import os
from xmind2testcase.utils import get_xmind_testcase_list, get_absolute_path

"""
Convert XMind fie to Zentao testcase csv file 

Zentao official document about import CSV testcase file: https://www.zentao.net/book/zentaopmshelp/243.mhtml 
"""


def xmind_to_zentao_csv_file(xmind_file):
    """Convert XMind file to a zentao csv file"""
    xmind_file = get_absolute_path(xmind_file)
    logging.info('Start converting XMind file(%s) to zentao file...', xmind_file)
    testcases = get_xmind_testcase_list(xmind_file)
    # fileheader = ["所属模块", "用例标题", "前置条件", "步骤", "预期", "关键词", "优先级", "用例类型", "适用阶段"]
    fileheader = ["ID", "一级模块", "二级模块", "用例名称", "优先级", "用例类型", "前置条件", "步骤描述", "预期结果", "备注", "维护人", "需求描述"]
    zentao_testcase_rows = [fileheader]
    for testcase in testcases:
        row = gen_a_testcase_row(testcase)
        zentao_testcase_rows.append(row)
    zentao_file = xmind_file[:-6] + '.csv'
    if os.path.exists(zentao_file):
        os.remove(zentao_file)
    with open(zentao_file, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerows(zentao_testcase_rows)
        logging.info('Convert XMind file(%s) to a zentao csv file(%s) successfully!', xmind_file, zentao_file)
    return zentao_file

def xmind_to_ones_xlsx_file(xmind_file): #xmind导出禅道用例模板的xlsx格式
    """Convert XMind file to a zentao xlsx file"""
    xmind_file = get_absolute_path(xmind_file)
    logging.info('Start converting XMind file(%s) to zentao file...', xmind_file)
    testcases = get_xmind_testcase_list(xmind_file)

    fileheader = ["ID", "一级模块", "二级模块", "用例名称", "优先级", "用例类型", "前置条件", "步骤描述", "预期结果", "备注", "维护人", "需求描述"]
    zentao_testcase_rows = []
    for testcase in testcases:
        row = gen_a_testcase_row(testcase)
        zentao_testcase_rows.append(row)

    zentao_file = xmind_file[:-6] + '.xlsx'
    df = pd.DataFrame(data=zentao_testcase_rows, columns=fileheader) #构造数据
    df.to_excel(zentao_file, index=False)  #写入文件，设置不需要索引
    logging.info('Convert XMind file(%s) to a zentao xlsx file(%s) successfully!', xmind_file, zentao_file)
    return zentao_file


def gen_a_testcase_row(testcase_dict):
    case_id = ''
    case_module0 = ''
    case_module = gen_case_module(testcase_dict['suite'])
    case_title = testcase_dict['name']
    case_precontion = testcase_dict['preconditions']
    case_step, case_expected_result = gen_case_step_and_expected_result(testcase_dict['steps'])
    case_priority = gen_case_priority(testcase_dict['importance'])
    case_type = gen_case_type(testcase_dict['execution_type'])
    case_tips = ''
    case_creator = '齐通通'
    case_requirement = ''
    row = [case_id, case_module0, case_module, case_title, case_priority, case_type, case_precontion, case_step,
           case_expected_result, case_tips, case_creator, case_requirement]
    return row


def gen_case_module(module_name):
    if module_name:
        module_name = module_name.replace('（', '(')
        module_name = module_name.replace('）', ')')
    else:
        module_name = '/'
    return module_name


def gen_case_step_and_expected_result(steps):
    case_step = ''
    case_expected_result = ''

    for step_dict in steps:
        case_step += str(step_dict['step_number']) + '. ' + step_dict['actions'].replace('\n', '').strip() + '\n'
        case_expected_result += str(step_dict['step_number']) + '. ' + \
                                step_dict['expectedresults'].replace('\n', '').strip() + '\n' \
            if step_dict.get('expectedresults', '') else ''

    return case_step, case_expected_result


def gen_case_priority(priority):
    mapping = {1: 'P1', 2: 'P2', 3: 'P3', 4: 'P4'}
    # mapping = {1: '高', 2: '中', 3: '低'}
    if priority in mapping.keys():
        return mapping[priority]
    else:
        return 'P2'


def gen_case_type(case_type):
    # mapping = {1: '手动', 2: '自动'}
    mapping = {1: '功能测试', 2: '接口测试'}
    if case_type in mapping.keys():
        return mapping[case_type]
    else:
        return '功能测试'


if __name__ == '__main__':
    xmind_file = '../docs/zentao_testcase_template.xmind'
    zentao_csv_file = xmind_to_zentao_csv_file(xmind_file)
    print('Conver the xmind file to a zentao csv file succssfully: %s', zentao_csv_file)
