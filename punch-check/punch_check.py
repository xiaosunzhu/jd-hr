# coding=utf-8

import ConfigParser
import sys
from datetime import date, time

import xlrd

from sheet_read_write import read_str_cell, read_cell_type, read_int_cell, read_date_cells, read_time_cells, \
    write_by_date_sheet_row, write_final_sheet_row, write_details_sheet_row, outputData, write_no_plan_sheet_row, \
    MSG_NOT_PUNCH_IN, MSG_PUNCH_IN_LATE, MSG_NOT_PUNCH_OUT, MSG_PUNCH_OUT_EARLY, MSG_NOT_PUNCH, write_details_plan_col
from work_def import Person, WorkDay, Punch, get_date_time, PlanType, FLOAT_TYPE, is_same_time_punch, RestDay


__author__ = 'yijun.sun'

reload(sys)
sys.setdefaultencoding("utf-8")

SYSTEM_ENCODING = 'GBK'
year = 2015
month = 2


def encode_str(string):
    return string.encode(SYSTEM_ENCODING)


PLAN_DEPARTMENT_MAP = {}

print(encode_str('Copyright 2015 yijun.sun'))
print(encode_str('Version: 0.0.3'))

try:
    tableConfig = ConfigParser.ConfigParser()
    try:
        with open(encode_str('config\\表格配置.ini'), 'r') as cfg_file:
            tableConfig.readfp(cfg_file)
    except IOError, e:
        print(encode_str('无法打开表格配置！'))
        raise
    try:
        planTableNameCol = int(tableConfig.get(encode_str('排班表'), encode_str('姓名列')).strip()) - 1
        planTableNameStartRow = int(tableConfig.get(encode_str('排班表'), encode_str('姓名起始行')).strip()) - 1
        planTableDepartmentCol = int(tableConfig.get(encode_str('排班表'), encode_str('部门列')).strip()) - 1
        planTableDateRow = int(tableConfig.get(encode_str('排班表'), encode_str('日期行')).strip()) - 1
        planTableDateStartCol = int(tableConfig.get(encode_str('排班表'), encode_str('日期起始列')).strip()) - 1
        planSheetIndex = int(tableConfig.get(encode_str('排班表'), encode_str('Sheet')).strip()) - 1
        punchDepartmentCol = int(tableConfig.get(encode_str('打卡表'), encode_str('部门列')).strip()) - 1
        punchTableNameCol = int(tableConfig.get(encode_str('打卡表'), encode_str('姓名列')).strip()) - 1
        punchDateCol = int(tableConfig.get(encode_str('打卡表'), encode_str('日期列')).strip()) - 1
        punchTimeCol = int(tableConfig.get(encode_str('打卡表'), encode_str('时间列')).strip()) - 1
        punchTypeCol = int(tableConfig.get(encode_str('打卡表'), encode_str('类型列')).strip()) - 1
        punchNameStartRow = int(tableConfig.get(encode_str('打卡表'), encode_str('姓名起始行')).strip()) - 1
        punchSheetIndex = int(tableConfig.get(encode_str('打卡表'), encode_str('Sheet')).strip()) - 1
    except Exception, e:
        print(encode_str('表格配置 格式非法！'))
        raise

    globalPlanSection = '请假'

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
                if departmentDecode != globalPlanSection:
                    needWork = True
                PLAN_DEPARTMENT_MAP[departmentDecode][planCodeString] = PlanType(planCodeString, describe,
                                                                                 beginTime,
                                                                                 endTime,
                                                                                 acrossDay, needWork)

    except Exception, e:
        print(encode_str('排班代码配置 格式非法！'))
        raise

    # planFilePath = raw_input('排班表：'.encode(SYSTEM_ENCODING))
    # punchFilePath = raw_input('打卡表：'.encode(SYSTEM_ENCODING))
    # planFilePath = planFilePath.replace('"', "")
    # punchFilePath = punchFilePath.replace('"', "")

    planFilePath = encode_str('resources\\2月28运输排班汇总表（双） .xlsx')
    punchFilePath = encode_str('resources\\打卡记录.xls')

    startDateNum = 1
    endDateNum = 1
    try:
        planData = xlrd.open_workbook(planFilePath)
    except IOError, e:
        print(encode_str('无法打开排班表！\n文件路径：') + planFilePath)
        raise
    except xlrd.XLRDError, e:
        print(encode_str('排班表是已损坏的excel文件！\n文件路径：') + planFilePath)
        raise
    planSheet = planData.sheets()[planSheetIndex]
    personMap = {}
    globalPlanTimeMap = PLAN_DEPARTMENT_MAP.get(globalPlanSection)
    for row in range(planTableNameStartRow, planSheet.nrows):
        name = read_str_cell(planSheet, row, planTableNameCol)
        department = read_str_cell(planSheet, row, planTableDepartmentCol).strip()
        planTimeMap = PLAN_DEPARTMENT_MAP.get(department)
        if planTimeMap is None:
            continue
        if name.strip() == '':
            continue
        if name not in personMap.keys():
            personMap[name] = Person(name, department)
        if not planTimeMap:
            continue
        colNum = planTableDateStartCol
        while read_cell_type(planSheet, planTableDateRow, colNum) is FLOAT_TYPE:
            dateTemp = read_int_cell(planSheet, planTableDateRow, colNum)
            if dateTemp > endDateNum:
                endDateNum = dateTemp
            planType = read_str_cell(planSheet, row, colNum)
            if planType.strip() == '':
                colNum += 1
                continue
            planWork = planTimeMap.get(planType)
            if not planWork:
                planWork = globalPlanTimeMap.get(planType)
                if not planWork:
                    colNum += 1
                    continue
            if planWork.needWork:
                workPlan = WorkDay(date(year, month, dateTemp), planWork)
            else:
                workPlan = RestDay(date(year, month, dateTemp), planWork)
            personMap[name].add_day_plan(workPlan)
            colNum += 1

    detailsOutputRow = 1
    noPlanOutputRow = 1

    try:
        punchData = xlrd.open_workbook(punchFilePath)
    except IOError, e:
        print(encode_str('无法打开打卡表！\n文件路径：') + punchFilePath)
        raise
    except xlrd.XLRDError, e:
        print(encode_str('打卡表是已损坏的excel文件！\n文件路径：') + punchFilePath)
        raise
    punchSheet = punchData.sheets()[punchSheetIndex]
    processedNoPlanName = {}
    for row in range(punchNameStartRow, punchSheet.nrows):
        name = read_str_cell(punchSheet, row, punchTableNameCol)
        splits = name.split(' ')
        name = splits[len(splits) - 1]
        department = read_str_cell(punchSheet, row, punchDepartmentCol)
        currentDate = read_date_cells(punchSheet, punchData.datemode, row, punchDateCol)
        currentTime = read_time_cells(punchSheet, punchData.datemode, row, punchTimeCol)
        punchDatetime = get_date_time(currentDate, currentTime)
        punchType = read_str_cell(punchSheet, row, punchTypeCol)
        if name not in personMap.keys():
            if name not in processedNoPlanName.keys():
                noPlanOutputRow = write_no_plan_sheet_row(noPlanOutputRow, name,
                                                          department, detailsOutputRow + 1)
                processedNoPlanName[name] = noPlanOutputRow
            detailsOutputRow = write_details_sheet_row(detailsOutputRow, name, department,
                                                       punchDatetime, punchType,
                                                       processedNoPlanName[name], no_plan_sheet=True)
            continue
        person = personMap[name]
        person.add_punch(Punch(punchType, punchDatetime))

    for person in personMap.values():
        indexOfPunch = 0
        finishPersonPunchCheck = False
        for dateNum in range(startDateNum, endDateNum + 1):
            currentDate = date(year, month, dateNum)
            work = person.workDays.get(currentDate)
            if not work:
                continue
            if indexOfPunch >= len(person.punches):
                break
            while work.is_before_work_uncertain_time(person.punches[indexOfPunch]):
                indexOfPunch += 1
                if indexOfPunch >= len(person.punches):
                    break
            if indexOfPunch >= len(person.punches):
                break
            while work.is_before_work_valid_time(person.punches[indexOfPunch]):
                work.uncertain_punch_in(person.punches[indexOfPunch])
                person.punches[indexOfPunch].processed = True
                indexOfPunch += 1
                if indexOfPunch >= len(person.punches):
                    break
            if indexOfPunch >= len(person.punches):
                break
            while not work.is_after_work_valid_time(person.punches[indexOfPunch]):
                work.punch(person.punches[indexOfPunch])
                person.punches[indexOfPunch].processed = True
                indexOfPunch += 1
                if indexOfPunch >= len(person.punches):
                    break
            uncertainCount = 0
            if indexOfPunch >= len(person.punches):
                break
            while not work.is_after_work_uncertain_time(person.punches[indexOfPunch]):
                work.uncertain_punch_out(person.punches[indexOfPunch])
                person.punches[indexOfPunch].processed = True
                uncertainCount += 1
                indexOfPunch += 1
                if indexOfPunch >= len(person.punches):
                    break
            indexOfPunch -= uncertainCount

    nameSorted = sorted(personMap.keys())
    finalOutputRow = 1
    byDateOutputRow = 1

    for name in nameSorted:
        person = personMap[name]
        for dateNum in range(startDateNum, endDateNum + 1):
            currentDate = date(year, month, dateNum)
            work = person.workDays.get(currentDate)
            rest = person.restDays.get(currentDate)
            beforeDayWork = None
            nextDayWork = None
            if dateNum != startDateNum:
                beforeDayWork = person.workDays.get(date(year, month, dateNum - 1))
            if dateNum != endDateNum:
                nextDayWork = person.workDays.get(date(year, month, dateNum + 1))

            if rest:
                if not rest.haveOutput:
                    lastRestDateNum = dateNum
                    rest.mark_output()
                    while lastRestDateNum < endDateNum and person.restDays.get(date(year, month, lastRestDateNum + 1)):
                        if person.restDays.get(date(year, month, lastRestDateNum + 1)).plan == rest.plan:
                            person.restDays.get(date(year, month, lastRestDateNum + 1)).mark_output()
                            lastRestDateNum += 1
                        else:
                            break
                    lastRest = person.restDays.get(date(year, month, lastRestDateNum))
                    finalOutputRow = write_final_sheet_row(finalOutputRow, person.name,
                                                           person.department,
                                                           rest.get_plan_begin_datetime(),
                                                           lastRest.get_plan_end_datetime(),
                                                           rest.plan.describe.decode(SYSTEM_ENCODING), None)
                continue

            if not work:
                continue
            workDate = work.get_work_date()
            planType = work.get_plan_type()

            # 补充确定先前不确定的打卡记录
            if not work.have_punch_in() and work.uncertainPunchInList:
                work.punch(work.uncertainPunchInList[0])
            if not work.have_punch_out() and len(work.uncertainPunchOutList) > 0:
                uncertainPunchOutFirst = work.uncertainPunchOutList[0]
                uncertainPunchOutLast = work.uncertainPunchOutList[
                    len(work.uncertainPunchOutList) - 1]
                if not nextDayWork:
                    work.punch(uncertainPunchOutLast)
                elif nextDayWork.have_punch_in() or \
                                (uncertainPunchOutFirst.punchDatetime - work.get_plan_end_datetime()).seconds <= (
                                    (nextDayWork.get_plan_begin_datetime() - work.get_plan_end_datetime()).seconds / 2):
                    uncertainPunchOut = uncertainPunchOutFirst
                    for punchIn in work.uncertainPunchOutList:
                        if is_same_time_punch(uncertainPunchOut, punchIn):
                            uncertainPunchOut = punchIn
                        else:
                            break
                    work.punch(uncertainPunchOut)
                    nextDayWork.remove_processed_uncertain_punch_in(uncertainPunchOut.punchDatetime)
                    # 补充确定先前不确定的打卡记录

            exceptionMsg = ''
            if not work.have_punch_in():
                exceptionMsg += MSG_NOT_PUNCH_IN + ' / '
                finalOutputRow = write_final_sheet_row(finalOutputRow, person.name,
                                                       person.department,
                                                       work.get_plan_begin_datetime(),
                                                       work.get_plan_begin_datetime(),
                                                       MSG_NOT_PUNCH, byDateOutputRow + 1)
            elif work.is_punch_in_late():
                exceptionMsg += MSG_PUNCH_IN_LATE + ' / '
            if dateNum != endDateNum and not work.have_punch_out():
                exceptionMsg += MSG_NOT_PUNCH_OUT + ' / '
                finalOutputRow = write_final_sheet_row(finalOutputRow, person.name,
                                                       person.department,
                                                       work.get_plan_end_datetime(),
                                                       work.get_plan_end_datetime(),
                                                       MSG_NOT_PUNCH, byDateOutputRow + 1)
            elif work.is_punch_out_early():
                exceptionMsg += MSG_PUNCH_OUT_EARLY + ' / '
            if work.have_punch_in() and not work.is_punch_in_late() and work.have_punch_out() and not work.is_punch_out_early():
                pass
            detailsStartRow = detailsOutputRow + 1
            detailsLocateRow = None
            if exceptionMsg:
                exceptionMsg = exceptionMsg[:len(exceptionMsg) - 3]
                if beforeDayWork:
                    yesterdayPlan = beforeDayWork.planWork
                else:
                    yesterdayPlan = None
                todayPlan = work.planWork
                if nextDayWork:
                    tomorrowPlan = nextDayWork.planWork
                else:
                    tomorrowPlan = None
                yesterdayStartRow = None
                yesterdayEndRow = None
                todayStartRow = None
                todayEndRow = None
                tomorrowStartRow = None
                tomorrowEndRow = None
                for punch in person.punches:
                    if not punch.outputToDetails and (punch.punchDatetime.day == dateNum - 1 or
                                                              punch.punchDatetime.day == dateNum or punch.punchDatetime.day == dateNum + 1):
                        if not detailsLocateRow and punch.punchDatetime.day == dateNum:
                            detailsLocateRow = detailsOutputRow + 1
                        if not detailsLocateRow and punch.punchDatetime.day == dateNum + 1:
                            detailsLocateRow = detailsOutputRow
                        if punch.punchDatetime.day == dateNum - 1:
                            if not yesterdayStartRow:
                                yesterdayStartRow = detailsOutputRow
                            yesterdayEndRow = detailsOutputRow
                        elif punch.punchDatetime.day == dateNum:
                            if not todayStartRow:
                                todayStartRow = detailsOutputRow
                            todayEndRow = detailsOutputRow
                        elif punch.punchDatetime.day == dateNum + 1:
                            if not tomorrowStartRow:
                                tomorrowStartRow = detailsOutputRow
                            tomorrowEndRow = detailsOutputRow
                        detailsOutputRow = write_details_sheet_row(detailsOutputRow, person.name,
                                                                   person.department,
                                                                   punch.punchDatetime,
                                                                   punch.punchType,
                                                                   byDateOutputRow + 1,
                                                                   time_info_sheet=True)
                        punch.outputToDetails = True
                if yesterdayPlan and yesterdayStartRow:
                    write_details_plan_col(yesterdayStartRow, yesterdayEndRow, yesterdayPlan)
                if todayPlan and todayStartRow:
                    write_details_plan_col(todayStartRow, todayEndRow, todayPlan)
                if tomorrowPlan and tomorrowStartRow:
                    write_details_plan_col(tomorrowStartRow, tomorrowEndRow, tomorrowPlan)
            if not detailsLocateRow:
                detailsLocateRow = detailsStartRow
            byDateOutputRow = write_by_date_sheet_row(byDateOutputRow, person.name,
                                                      workDate, work.get_punch_in_datetime(),
                                                      work.get_punch_out_datetime(), planType,
                                                      exceptionMsg, detailsLocateRow)
    try:
        outputData.save(encode_str('排班打卡比对_' + str(year) + '年' + str(month) + '月.xls'))
        print(encode_str('处理完毕'))
    except IOError, e:
        print(encode_str('无法写入表格文件。请确认已关闭该文件并且有操作权限！'))
        raise
except Exception, e:
    print('\n' + encode_str('程序异常！') + ' ' + e.message)
    raise
finally:
    raw_input(encode_str('键入回车退出程序'))
