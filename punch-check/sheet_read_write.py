# coding=utf-8

from datetime import time, date

import xlrd
import xlwt
from xlwt import Style

from work_def import STRING_TYPE, FLOAT_TYPE


__author__ = 'yijun.sun'

MSG_NOT_PUNCH = '未打卡'
MSG_NOT_PUNCH_IN = '上班' + MSG_NOT_PUNCH
MSG_NOT_PUNCH_OUT = '下班' + MSG_NOT_PUNCH
MSG_PUNCH_IN_LATE = '迟到'
MSG_PUNCH_OUT_EARLY = '早退'


def read_cell_type(sheet, row, col):
    data = sheet.cell(row, col).value
    return type(data)


def read_int_cell(sheet, row, col):
    data = sheet.cell(row, col).value
    return int(data)


def read_str_cell(sheet, row, col, not_float=True):
    data = sheet.cell(row, col).value
    if type(data) is STRING_TYPE:
        return data.strip()
    elif type(data) is FLOAT_TYPE:
        if not_float:
            return str(int(data))
    else:
        return str(data)


def read_date_cells(sheet, date_mode, date_row, dat_col):
    date_tuple = xlrd.xldate_as_tuple(sheet.cell(date_row, dat_col).value, date_mode)
    return date(*date_tuple[:3])


def read_time_cells(sheet, date_mode, time_row, time_col):
    time_tuple = xlrd.xldate_as_tuple(sheet.cell(time_row, time_col).value, date_mode)
    return time(*time_tuple[3:])


YELLOW_BG_PATTERN = xlwt.Pattern()
YELLOW_BG_PATTERN.pattern = xlwt.Pattern.SOLID_PATTERN
YELLOW_BG_PATTERN.pattern_fore_colour = 5  # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...

MAGENTA_BG_PATTERN = xlwt.Pattern()
MAGENTA_BG_PATTERN.pattern = xlwt.Pattern.SOLID_PATTERN
MAGENTA_BG_PATTERN.pattern_fore_colour = 6  # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...


def write_details_sheet_row(row, name, department, punch_datetime, punch_type,
                            link_exception_sheet_row):
    outputDetailsSheet.write(row, 0, name)
    outputDetailsSheet.write(row, 1, department)
    outputDetailsSheet.write(row, 2, xlwt.Formula(
        'HYPERLINK("#timeInfo!A' + str(link_exception_sheet_row) + '","' + str(punch_datetime)[
                                                                           0:10] + '")'))
    outputDetailsSheet.write(row, 3, punch_datetime.strftime('%H:%M:%S'))
    outputDetailsSheet.write(row, 4, punch_type)
    return row + 1


def write_by_date_sheet_row(row, name, date, punch_in_datetime, punch_out_datetime, plan, msg,
                            link_details_sheet_row):
    style = Style.default_style
    origin_pattern = style.pattern
    if msg:
        if MSG_NOT_PUNCH in msg:
            style.pattern = MAGENTA_BG_PATTERN
        else:
            style.pattern = YELLOW_BG_PATTERN
        outputByDateSheet.write(row, 6, msg)
    outputByDateSheet.write(row, 0, name)
    outputByDateSheet.write(row, 1, xlwt.Formula(
        'HYPERLINK("#details!A' + str(link_details_sheet_row) + '","' + str(date)[0:10] + '")'))
    if punch_in_datetime:
        outputByDateSheet.write(row, 2, str(punch_in_datetime.time()))
    else:
        outputByDateSheet.write(row, 2, '')
    if punch_out_datetime:
        outputByDateSheet.write(row, 3, str(punch_out_datetime.time()))
    else:
        outputByDateSheet.write(row, 3, '')
    style.num_format_str = 'h:mm'
    rowNum = str(row + 1)
    outputByDateSheet.write(row, 4, xlwt.Formula(
        #IF(D?*C?,IF(D?<C?,D?+"24:00:00",D?)-C?),"")
        'IF(D' + rowNum + '*C' + rowNum + ',IF(D' + rowNum + '<C' + rowNum + ',D' + rowNum +
        '+"24:00:00",D' + rowNum + ')-C' + str(row + 1) + ',"")'))
    style.num_format_str = 'General'
    outputByDateSheet.write(row, 5, str(plan), style)
    style.pattern = origin_pattern
    return row + 1


def write_final_sheet_row(row, name, department, leave_start, leave_end, type, link_exception_row):
    outputFinalSheet.write(row, 0, row)
    outputFinalSheet.write(row, 1, name)
    outputFinalSheet.write(row, 2, department)
    outputFinalSheet.write(row, 4, str(leave_start))
    outputFinalSheet.write(row, 5, xlwt.Formula(
        'HYPERLINK("#timeInfo!A' + str(link_exception_row) + '","' + str(leave_end) + '")'))
    outputFinalSheet.write(row, 6, type)
    return row + 1


def write_no_plan_sheet_row(row, name, department):
    outputNoPlanSheet.write(row, 0, row)
    outputNoPlanSheet.write(row, 1, name)
    outputNoPlanSheet.write(row, 2, department)
    return row + 1


outputData = xlwt.Workbook(encoding='utf-8', style_compression=0)

outputFinalSheet = outputData.add_sheet('考勤异常')
outputFinalSheet.col(0).width = 256 * 6
outputFinalSheet.col(1).width = 256 * 12
outputFinalSheet.col(2).width = 256 * 16
outputFinalSheet.col(3).width = 256 * 18
outputFinalSheet.col(4).width = 256 * 24
outputFinalSheet.col(5).width = 256 * 24
outputFinalSheet.col(6).width = 256 * 16
outputFinalSheet.col(7).width = 256 * 14
outputFinalSheet.col(8).width = 256 * 15
outputFinalSheet.col(9).width = 256 * 14
outputFinalSheet.write(0, 0, '序号')
outputFinalSheet.write(0, 1, '姓名')
outputFinalSheet.write(0, 2, '部门')
outputFinalSheet.write(0, 3, '岗位')
outputFinalSheet.write(0, 4, '起假时间')
outputFinalSheet.write(0, 5, '截止时间')
outputFinalSheet.write(0, 6, '假别')
outputFinalSheet.write(0, 7, '时间段')
outputFinalSheet.write(0, 8, '请假时间（小时）')
outputFinalSheet.write(0, 9, '状态')

outputByDateSheet = outputData.add_sheet('timeInfo')
outputByDateSheet.col(0).width = 256 * 12
outputByDateSheet.col(1).width = 256 * 15
outputByDateSheet.col(2).width = 256 * 15
outputByDateSheet.col(3).width = 256 * 15
outputByDateSheet.col(4).width = 256 * 12
outputByDateSheet.col(5).width = 256 * 20
outputByDateSheet.col(6).width = 256 * 30
outputByDateSheet.write(0, 0, '姓名')
outputByDateSheet.write(0, 1, '日期')
outputByDateSheet.write(0, 2, '上班卡时间')
outputByDateSheet.write(0, 3, '下班卡时间')
outputByDateSheet.write(0, 4, '在班时间')
outputByDateSheet.write(0, 5, '排班')
outputByDateSheet.write(0, 6, '异常信息')

outputDetailsSheet = outputData.add_sheet('details')

outputNoPlanSheet = outputData.add_sheet('未排班人员')
outputNoPlanSheet.col(0).width = 256 * 6
outputNoPlanSheet.col(1).width = 256 * 12
outputNoPlanSheet.col(2).width = 256 * 16
outputNoPlanSheet.write(0, 0, '序号')
outputNoPlanSheet.write(0, 1, '姓名')
outputNoPlanSheet.write(0, 2, '部门')
