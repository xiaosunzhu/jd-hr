# coding=utf-8
from time import sleep
import traceback

from check_update import request_to_github


__author__ = 'yijun.sun'

import sys

from base import *

reload(sys)
sys.setdefaultencoding("utf-8")

print(encode_str('Copyright 2015 yijun.sun'))
print(encode_str('Version: ' + CURRENT_VERSION))
print('')


def get_valid_part(name_string):
    splits = name_string.split(' ')
    name_string = splits[len(splits) - 1]
    splits = name_string.split('（')
    name_string = splits[0]
    splits = name_string.split('(')
    return splits[0].strip()


try:
    result = request_to_github()
    if result:
        print(encode_str('*** 如需更新软件，请运行punch_update.exe ***'))
        print('')

    from configs import *
    from sheet_read_write import *  # 需要先导入configs模块再导入sheet_read_write
    from work_def import *

    planFilePath = raw_input(encode_str('排班表：'))
    punchFilePath = raw_input(encode_str('打卡表：'))
    planFilePath = planFilePath.replace('"', "")
    punchFilePath = punchFilePath.replace('"', "")

    # planFilePath = encode_str('resources\\3月排班表_仓库.xlsx')
    # punchFilePath = encode_str('resources\\3月指纹_仓库.xls')

    startDateNum = 1
    dateCount = 0
    dates = []
    identitySorted = []
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
    restPlanTimeMap = PLAN_DEPARTMENT_MAP.get(restPlanSection)
    globalPlanTimeMap = PLAN_DEPARTMENT_MAP.get(globalPlanSection)
    haveSetDates = False
    notSetCode = []
    repeatedPerson = {}
    for row in range(planTablePersonStartRow, planSheet.nrows):
        identity = read_str_cell(planSheet, row, planTableIdentityCol).strip()
        name = read_str_cell(planSheet, row, planTableNameCol).strip()
        department = read_str_cell(planSheet, row, planTableDepartmentCol).strip()
        planTimeMap = None
        if not useGlobalPan:
            planTimeMap = PLAN_DEPARTMENT_MAP.get(department)
        else:
            planTimeMap = globalPlanTimeMap
        if planTimeMap is None:
            continue
        if identity.strip() == '':
            continue
        if identity not in personMap.keys():
            personMap[identity] = Person(identity, name, department)
            identitySorted.append(identity)
        if not planTimeMap:
            continue
        colNum = planTableDateStartCol
        currentMonth = startMonth
        lastDateNum = 0
        while len(planSheet.row(planTableDateRow)) > colNum and read_cell_type(planSheet, planTableDateRow,
                                                                               colNum) is FLOAT_TYPE:
            dateTempNum = read_int_cell(planSheet, planTableDateRow, colNum)
            if lastDateNum > dateTempNum:
                currentMonth += 1
            dateTemp = date(year, currentMonth, dateTempNum)
            lastDateNum = dateTempNum
            if not haveSetDates and dateTemp not in dates:
                dates.append(dateTemp)
            planType = read_str_cell(planSheet, row, colNum)
            if planType.strip() == '':
                colNum += 1
                continue
            planType = planType.upper().strip()
            planWork = planTimeMap.get(planType)
            if not planWork:
                planWork = restPlanTimeMap.get(planType)
                if not planWork:
                    colNum += 1
                    if planType not in notSetCode:
                        notSetCode.append(planType)
                    continue
            if planWork.needWork:
                workPlan = WorkDay(dateTemp, planWork)
            else:
                workPlan = RestDay(dateTemp, planWork)
            if personMap[identity].have_different_planed(dateTemp,
                                                         planType) and identity not in repeatedPerson.keys():
                repeatedPerson[identity] = name
            personMap[identity].add_day_plan(workPlan)
            colNum += 1
        haveSetDates = True
    if notSetCode:
        notSetCodeStr = ''
        for code in notSetCode:
            notSetCodeStr += ' ' + code
        print(encode_str('\n排班表中发现未配置的排班代码:' + notSetCodeStr + '\n'))
    if repeatedPerson:
        repeatedPersonStr = ''
        for identity in repeatedPerson.keys():
            repeatedPersonStr += '编号：' + identity + '，姓名：' + repeatedPerson[identity] + '\n'
        raise SelfException(encode_str('\n排班表中发现重复人员:\n' + repeatedPersonStr))

    dates.sort()
    oneDate = dates[0]
    lastDate = dates[len(dates) - 1]
    try:
        fromDateStr = raw_input(encode_str('排班起始日期为' + str(oneDate) + '，回车确认或输入起始日期：'))
        fromDate = oneDate
        if fromDateStr:
            fromDate = parse_str_to_date(fromDateStr)

        endDateStr = raw_input(encode_str('排班截止日期为' + str(lastDate) + '，回车确认或输入截止日期：'))
        endDate = lastDate
        if endDateStr:
            endDate = parse_str_to_date(endDateStr)
    except Exception, e:
        raise SelfException(encode_str('输入日期格式错误。格式为：年-月-日，如：2015-3-18'))
    print(encode_str('设定的处理时间段为：' + str(fromDate) + ' - ' + str(endDate)))
    print('')
    print(encode_str('请稍后......'))

    dateIndex = 0
    while oneDate <= lastDate:
        if (oneDate < fromDate or oneDate > endDate) and oneDate in dates:
            dates.remove(oneDate)
            oneDate = oneDate + timedelta(days=1)
            continue
        if oneDate not in dates:
            dates.append(oneDate)
        oneDate = oneDate + timedelta(days=1)
    dates.sort()

    detailsOutputRow = 1
    noPlanOutputRow = 1
    countOutputRow = 2

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
    row = 0
    dateOrDatetimeCellValid = False
    timeCellValid = False
    try:
        for row in range(punchPersonStartRow, punchSheet.nrows):
            dateOrDatetimeCellValid = False
            timeCellValid = False
            identity = read_str_cell(punchSheet, row, punchTableIdentityCol)
            name = read_str_cell(punchSheet, row, punchTableNameCol)
            identity = get_valid_part(identity)
            name = get_valid_part(name)
            department = read_str_cell(punchSheet, row, punchDepartmentCol)
            punchDatetime = None
            if punchSheetDatetimeNotSplit:
                punchDatetime = read_datetime_cells(punchSheet, punchData.datemode, row,
                                                    punchDateCol)
                dateOrDatetimeCellValid = True
            else:
                currentDate = read_date_cells(punchSheet, punchData.datemode, row, punchDateCol)
                dateOrDatetimeCellValid = True
                currentTime = read_time_cells(punchSheet, punchData.datemode, row, punchTimeCol)
                timeCellValid = True
                punchDatetime = get_date_time(currentDate, currentTime)
            punchType = read_str_cell(punchSheet, row, punchTypeCol)
            if identity not in personMap.keys():
                if identity not in processedNoPlanName.keys():
                    noPlanOutputRow = write_no_plan_sheet_row(noPlanOutputRow, identity, name,
                                                              department, detailsOutputRow + 1)
                    processedNoPlanName[identity] = noPlanOutputRow
                detailsOutputRow = write_details_sheet_row(detailsOutputRow, identity, name,
                                                           department,
                                                           punchDatetime, punchType,
                                                           processedNoPlanName[identity],
                                                           no_plan_sheet=True)
                continue
            person = personMap[identity]
            if person.name != name:
                raise SelfException(
                    encode_str('打卡表的人员编号和姓名与排班表不符！编号：' + str(
                        identity) + '，排班表姓名：' + person.name + '，打卡表姓名：' + name))
            person.add_punch(Punch(punchType, punchDatetime))
    except Exception, e:
        if isinstance(e, SelfException):
            raise e
        else:
            errColName = ''
            splitMsgHint = ''
            if punchSheetDatetimeNotSplit:
                splitMsgHint = '合并'
            else:
                splitMsgHint = '拆分'
            if dateOrDatetimeCellValid:
                errColName = '时间'
            elif punchSheetDatetimeNotSplit:
                errColName = '日期时间'
            else:
                errColName = '日期'
            raise SelfException(
                encode_str('打卡表采用日期时间' + splitMsgHint + '方式。第' + str(row + 1) + '行，' + errColName +
                           '列 格式错误'))

    for person in personMap.values():
        person.punches = sorted(person.punches,
                                cmp=lambda x, y: cmp(x.punchDatetime, y.punchDatetime))
        indexOfPunch = 0
        finishPersonPunchCheck = False
        for currentDate in dates:
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
                if work.have_punch_out() and (
                        not can_be_in_out_diff_punch_type(work.punchOut, person.punches[indexOfPunch])):
                    work.punch(person.punches[indexOfPunch])
                else:
                    work.uncertain_punch_out(person.punches[indexOfPunch])
                    person.punches[indexOfPunch].processed = True
                    uncertainCount += 1
                indexOfPunch += 1
                if indexOfPunch >= len(person.punches):
                    break
            indexOfPunch -= uncertainCount

    finalOutputRow = 1
    byDateOutputRow = 1

    for identity in identitySorted:
        person = personMap[identity]
        for index in range(0, len(dates)):
            currentDate = dates[index]
            work = person.workDays.get(currentDate)
            rest = person.restDays.get(currentDate)
            startDate = dates[0]
            endDate = dates[len(dates) - 1]
            beforeDayWork = None
            nextDayWork = None
            if index > 0:
                beforeDayWork = person.workDays.get(dates[index - 1])
            if index < len(dates) - 1:
                nextDayWork = person.workDays.get(dates[index + 1])

            if rest:
                if not rest.haveOutput and rest.get_plan_begin_datetime() and rest.get_plan_end_datetime():
                    lastRestDateIndex = index
                    rest.mark_output()
                    restDayCount = 1
                    while lastRestDateIndex < (len(dates) - 1) and \
                            person.restDays.get(dates[lastRestDateIndex + 1]):
                        if person.restDays.get(dates[lastRestDateIndex + 1]).plan == rest.plan:
                            person.restDays.get(dates[lastRestDateIndex + 1]).mark_output()
                            restDayCount += 1
                            lastRestDateIndex += 1
                        else:
                            break
                    lastRest = person.restDays.get(dates[lastRestDateIndex])
                    finalOutputRow = write_final_sheet_row(finalOutputRow, person.identity,
                                                           person.name,
                                                           person.department,
                                                           rest.get_plan_begin_datetime(),
                                                           lastRest.get_plan_end_datetime(),
                                                           rest.plan.describe.decode(
                                                               SYSTEM_ENCODING), None, restDayCount)
                continue

            if not work:
                continue
            workDate = work.get_work_date()
            planType = work.get_plan_type()

            # 补充确定先前不确定的打卡记录
            if work.needPunchIn and work.uncertainPunchInList and (
                        not work.have_punch_in() or work.is_punch_in_late()):
                firstUncertainPunchIn = work.uncertainPunchInList[0]
                mayBeEarlyPunchOut = None
                if work.is_punch_in_late():
                    # 本来是迟到，如果迟到时间和上班时间基本一致，而又离不确定时间不太远，那么不确定时间可以认定为上班打卡来消除迟到
                    # 如果迟到时间和上班时间不一致，就有可能用不确定时间代替原迟到时间，而原迟到时间可能判定为早退时间
                    if is_same_time(work.get_plan_begin_datetime(), work.punchIn.punchDatetime):
                        for uncertainPunchIn in work.uncertainPunchInList:
                            if not can_be_in_out_diff_punch_type(uncertainPunchIn, work.punchIn):
                                work.punch(uncertainPunchIn)
                                break
                    elif can_be_in_out_diff_punch_type(firstUncertainPunchIn, work.punchIn):
                        mayBeEarlyPunchOut = work.punchIn
                        work.punch(firstUncertainPunchIn)
                else:
                    work.punch(firstUncertainPunchIn)
                if mayBeEarlyPunchOut:
                    work.punch(mayBeEarlyPunchOut)
            if work.needPunchOut and work.uncertainPunchOutList and (
                        not work.have_punch_out() or work.is_punch_out_early()):
                uncertainPunchOutFirstGroup = work.uncertainPunchOutList[0]
                uncertainPunchOutLast = work.uncertainPunchOutList[
                    len(work.uncertainPunchOutList) - 1]
                haveMoreThanOneGroup = False
                for uncertainPunchOut in work.uncertainPunchOutList:
                    if can_be_in_out_diff_punch_type(uncertainPunchOut, uncertainPunchOutLast):
                        haveMoreThanOneGroup = True
                        uncertainPunchOutFirstGroup = uncertainPunchOut
                    else:
                        if not haveMoreThanOneGroup:
                            uncertainPunchOutFirstGroup = uncertainPunchOutLast
                        break
                if not nextDayWork or haveMoreThanOneGroup or (
                            nextDayWork.have_punch_in() and not nextDayWork.is_punch_in_late()) \
                        or (not nextDayWork.have_punch_in() and
                                    (
                                                uncertainPunchOutFirstGroup.punchDatetime - work.get_plan_end_datetime()).seconds
                                    <= (
                                            nextDayWork.get_plan_begin_datetime() - uncertainPunchOutFirstGroup.punchDatetime).seconds) \
                        or (nextDayWork.is_punch_in_late() and (can_be_in_out_diff_datetime(
                            work.uncertainPunchOutList[0].punchDatetime,
                            nextDayWork.get_plan_begin_datetime())
                                                                and (
                                        is_same_time(nextDayWork.get_plan_begin_datetime(),
                                                     nextDayWork.get_punch_in_datetime()) or (
                                                uncertainPunchOutFirstGroup.punchDatetime - work.get_plan_end_datetime()).seconds
                                        <= (
                                                    nextDayWork.get_punch_in_datetime() - uncertainPunchOutFirstGroup.punchDatetime).seconds))):
                    work.punch(uncertainPunchOutFirstGroup)
                    if nextDayWork:
                        nextDayWork.remove_processed_uncertain_punch_in(
                            uncertainPunchOutFirstGroup.punchDatetime)
                        # 补充确定先前不确定的打卡记录
            # 调整打卡确定是上班还是下班
            if not work.have_punch_out() and work.have_punch_in():
                if work.get_punch_in_datetime() > work.get_plan_begin_datetime() + timedelta(
                        seconds=(work.get_plan_end_datetime() - work.get_plan_begin_datetime()).seconds / 2):
                    work.punchOut = work.punchIn
                    work.havePunchOut = True
                    if work.get_punch_in_datetime() < work.get_plan_end_datetime():
                        work.punchOutEarly = True
                    work.havePunchIn = False
                    work.punchIn = None
                    work.punchInLate = False
            exceptionMsg = ''
            if work.needPunchIn and not work.have_punch_in() and 0 < index < (len(dates) - 1):
                exceptionMsg += MSG_NOT_PUNCH_IN + ' / '
                work.notPunchInRow = finalOutputRow
            elif work.needPunchIn and work.is_punch_in_late():
                exceptionMsg += MSG_PUNCH_IN_LATE + ' / '
            if index < (len(dates) - 1) \
                    and work.validEndDatetime < get_date_time(endDate, time(
                        6)) and work.needPunchOut and not work.have_punch_out():
                exceptionMsg += MSG_NOT_PUNCH_OUT + ' / '
                work.notPunchOutRow = finalOutputRow
            elif work.needPunchOut and work.is_punch_out_early():
                exceptionMsg += MSG_PUNCH_OUT_EARLY + ' / '
            if work.notPunchInRow and work.notPunchOutRow:
                finalOutputRow = write_final_sheet_row(finalOutputRow, person.identity, person.name,
                                                       person.department,
                                                       work.get_plan_begin_datetime(),
                                                       work.get_plan_end_datetime(),
                                                       MSG_NOT_PUNCH_BOTH, byDateOutputRow + 1)
            else:
                if work.countType:
                    person.countMap[work.countType][PERSON_COUNT_MAP_ARRAY_REAL_COUNT_INDEX] += 1
                if work.notPunchInRow:
                    finalOutputRow = write_final_sheet_row(finalOutputRow, person.identity, person.name,
                                                           person.department,
                                                           work.get_plan_begin_datetime(),
                                                           work.get_plan_begin_datetime(),
                                                           MSG_NOT_PUNCH, byDateOutputRow + 1)
                elif work.notPunchOutRow:
                    finalOutputRow = write_final_sheet_row(finalOutputRow, person.identity, person.name,
                                                           person.department,
                                                           work.get_plan_end_datetime(),
                                                           work.get_plan_end_datetime(),
                                                           MSG_NOT_PUNCH, byDateOutputRow + 1)
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
                    if not punch.notReal and not punch.outputToDetails and (
                                    (index > 0 and punch.punchDatetime.date() == dates[index - 1]) or
                                        punch.punchDatetime.date() == currentDate or
                                (index < (len(dates) - 1) and punch.punchDatetime.date() == dates[
                                        index + 1])):
                        if not detailsLocateRow and punch.punchDatetime.date() == currentDate:
                            detailsLocateRow = detailsOutputRow + 1
                        if not detailsLocateRow and punch.punchDatetime.date() > currentDate:
                            detailsLocateRow = detailsOutputRow
                        if punch.punchDatetime.date() < currentDate:
                            if not yesterdayStartRow:
                                yesterdayStartRow = detailsOutputRow
                            yesterdayEndRow = detailsOutputRow
                        elif punch.punchDatetime.date() == currentDate:
                            if not todayStartRow:
                                todayStartRow = detailsOutputRow
                            todayEndRow = detailsOutputRow
                        elif punch.punchDatetime.date() > currentDate:
                            if not tomorrowStartRow:
                                tomorrowStartRow = detailsOutputRow
                            tomorrowEndRow = detailsOutputRow
                        detailsOutputRow = write_details_sheet_row(detailsOutputRow,
                                                                   person.identity, person.name,
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
            haveDoubt = (work.is_punch_in_late() and work.punch_in_too_late()) \
                        or (work.is_punch_out_early() and work.punch_out_too_early())
            byDateOutputRow = write_by_date_sheet_row(byDateOutputRow, person.identity, person.name,
                                                      workDate, work.get_punch_in_datetime(),
                                                      work.get_punch_out_datetime(), planType,
                                                      exceptionMsg, haveDoubt, detailsLocateRow)
            # 连续异常加背景色
        for index in range(0, len(dates)):
            currentDate = dates[index]
            work = person.workDays.get(currentDate)
            if not work:
                continue
            if work.notPunchInRow and work.notPunchOutRow:  # 上下班均未打卡
                write_final_sheet_bg(*(work.notPunchInRow, work.notPunchOutRow))

        if NEED_COUNT_CODE_MAP:
            countOutputRow = write_count_sheet_row(countOutputRow, person.identity, person.name, person.department,
                                                   person.countMap)

    try:
        outputData.save(encode_str('排班打卡比对_' + str(year) + '年' + str(month) + '月.xls'))
        print(encode_str('处理完毕'))
    except IOError, e:
        print(encode_str('无法写入表格文件。请确认已关闭该文件并且有操作权限！'))
        raise
except Exception, e:
    print(encode_str('程序异常！ ') + str(e.message))
    print('')
    if not isinstance(e, SelfException):
        sleep(0.2)
        traceback.print_exc()
finally:
    sleep(0.6)
    raw_input(encode_str('键入回车退出程序'))
