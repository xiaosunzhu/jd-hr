# coding=utf-8

__author__ = 'SUNYIJUN'

import ConfigParser
from datetime import time

from base import *
from work_def import PlanType


baseConfig = ConfigParser.ConfigParser()
try:
    with open(encode_str('config\\数据配置.ini'), 'r') as cfg_file:
        baseConfig.readfp(cfg_file)
except IOError, e:
    print(encode_str('无法打开数据配置！'))
    raise
try:
    year = int(baseConfig.get(encode_str('Base'), encode_str('年')).strip())
    month = int(baseConfig.get(encode_str('Base'), encode_str('月')).strip())
    startMonth = int(baseConfig.get(encode_str('Base'), encode_str('起始月')).strip())
    useGlobalPan = int(baseConfig.get(encode_str('Swtiches'), encode_str('全局排班')).strip())
except Exception, e:
    print(encode_str('数据配置格式非法！'))
    raise

tableConfig = ConfigParser.ConfigParser()
try:
    with open(encode_str('config\\表格配置.ini'), 'r') as cfg_file:
        tableConfig.readfp(cfg_file)
except IOError, e:
    print(encode_str('无法打开表格配置！'))
    raise
try:
    planTableIdentityCol = int(tableConfig.get(encode_str('排班表'), encode_str('编号列')).strip()) - 1
    planTableNameCol = int(tableConfig.get(encode_str('排班表'), encode_str('姓名列')).strip()) - 1
    planTablePersonStartRow = int(tableConfig.get(encode_str('排班表'), encode_str('姓名起始行')).strip()) - 1
    planTableDepartmentCol = int(tableConfig.get(encode_str('排班表'), encode_str('部门列')).strip()) - 1
    planTableDateRow = int(tableConfig.get(encode_str('排班表'), encode_str('日期行')).strip()) - 1
    planTableDateStartCol = int(tableConfig.get(encode_str('排班表'), encode_str('日期起始列')).strip()) - 1
    planSheetIndex = int(tableConfig.get(encode_str('排班表'), encode_str('Sheet')).strip()) - 1
    punchDepartmentCol = int(tableConfig.get(encode_str('打卡表'), encode_str('部门列')).strip()) - 1
    punchTableIdentityCol = int(tableConfig.get(encode_str('打卡表'), encode_str('编号列')).strip()) - 1
    punchTableNameCol = int(tableConfig.get(encode_str('打卡表'), encode_str('姓名列')).strip()) - 1
    punchDateCol = int(tableConfig.get(encode_str('打卡表'), encode_str('日期列')).strip()) - 1
    punchTimeCol = int(tableConfig.get(encode_str('打卡表'), encode_str('时间列')).strip()) - 1
    punchTypeCol = int(tableConfig.get(encode_str('打卡表'), encode_str('类型列')).strip()) - 1
    punchPersonStartRow = int(tableConfig.get(encode_str('打卡表'), encode_str('姓名起始行')).strip()) - 1
    punchSheetIndex = int(tableConfig.get(encode_str('打卡表'), encode_str('Sheet')).strip()) - 1
except Exception, e:
    print(encode_str('表格配置格式非法！'))
    raise

restPlanSection = '请假'
globalPlanSection = 'Global'
notSetRestCode = '休'
notSetLeaveCode = '离'

PLAN_DEPARTMENT_MAP = {}

planCodeConfig = ConfigParser.ConfigParser()
try:
    with open(encode_str('config\\排班代码配置.ini'), 'r') as cfg_file:
        planCodeConfig.readfp(cfg_file)
except IOError, e:
    print(encode_str('无法打开排班代码配置！'))
    raise
try:
    for department in planCodeConfig.sections():
        departmentDecode = department.decode('GBK').encode('utf-8')
        PLAN_DEPARTMENT_MAP[departmentDecode] = {}
        departmentConfig = planCodeConfig.items(department)
        for planConfig in departmentConfig:
            planCodeString = planConfig[0].upper().decode('GBK').encode('utf-8')
            describe = None
            timeString = planConfig[1].decode('GBK').encode('utf-8')
            if ',' in planConfig[1]:
                describe = planConfig[1].split(',')[0]
                timeString = planConfig[1].split(',')[1]
            beginTime = None
            endTime = None
            acrossDay = False
            if '-' in timeString:
                timeStrings = timeString.split('-')
                beginString = timeStrings[0]
                if beginString == '':
                    continue
                endString = timeStrings[len(timeStrings) - 1]
                beginHour = int(beginString.split(':')[0])
                if beginHour == 24:
                    beginHour = 0
                beginTime = time(beginHour, int(beginString.split(':')[1]))
                endHour = int(endString.split(':')[0])
                if endHour == 24:
                    endHour = 0
                endTime = time(endHour, int(endString.split(':')[1]))
                if beginTime >= endTime:
                    acrossDay = True
            needWork = False
            if departmentDecode != restPlanSection:
                needWork = True
            PLAN_DEPARTMENT_MAP[departmentDecode][planCodeString] = PlanType(planCodeString,
                                                                             describe,
                                                                             beginTime,
                                                                             endTime,
                                                                             acrossDay, needWork)

except Exception, e:
    print(encode_str('排班代码配置格式非法！'))
    raise
