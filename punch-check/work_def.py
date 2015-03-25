# coding=utf-8

__author__ = 'yijun.sun'

from datetime import datetime, timedelta


FLOAT_TYPE = type(1.0)
STRING_TYPE = type('')


def enum(**enums):
    return type('Enum', (), enums)


class PlanType(object):
    def __init__(self, name, describe, begin_time=None, end_time=None, across_day=False,
                 need_work=True):
        self.name = name
        if not describe:
            self.describe = name
        else:
            self.describe = describe
        self.acrossDay = across_day
        self.beginTime = begin_time
        self.endTime = end_time
        if begin_time == None or end_time == None:
            self.needWork = False
        else:
            self.needWork = need_work

    def get_begin_time(self):
        return self.beginTime

    def get_end_time(self):
        return self.endTime

    def is_across_day(self):
        return self.acrossDay

    def __repr__(self):
        return self.name + ' (' + self.beginTime.strftime('%H:%M') + '-' + self.endTime.strftime(
            '%H:%M') + ')'


UNCERTAIN_WIN_HOURS_HALF = 4
NO_PLAN_VALID_EXPAND_HOURS = 10
NO_PLAN_UNCERTAIN_EXPAND_HOURS = 14
PUNCH_TYPE_DIFF_MIN_HOUR = 3
ONCE_PUNCH_DIFF_MAX_MINUTE = 50


class Person(object):
    def __init__(self, name, department):
        self.name = name
        self.department = department
        self.workDays = {}  # Map(date,WorkDay)
        self.restDays = {}  # Map(date,RestDay)
        self.punches = []  # Map(date,Punch[])

    def add_day_plan(self, day_plan):
        if isinstance(day_plan, RestDay):
            self.restDays[day_plan.workDate] = day_plan
        if isinstance(day_plan, WorkDay):
            self.workDays[day_plan.workDate] = day_plan
            yesterday = day_plan.workDate - timedelta(1)
            work_day_before = self.workDays.get(yesterday)
            current_begin = day_plan.get_plan_begin_datetime()

            if work_day_before:
                yesterday_end = work_day_before.get_plan_end_datetime()
                if yesterday_end == current_begin:
                    day_plan.set_valid_begin_datetime(current_begin)
                    day_plan.set_uncertain_punch_in_begin_datetime(current_begin)
                    day_plan.needPunchIn = False
                    day_plan.punch(Punch(PunchTypeKey.PunchIn, current_begin, True))
                    work_day_before.set_valid_end_datetime(current_begin)
                    work_day_before.set_uncertain_punch_out_end_datetime(current_begin)
                    work_day_before.needPunchOut = False
                    work_day_before.punch(Punch(PunchTypeKey.PunchOut, yesterday_end, True))
                else:
                    uncertain_begin = current_begin - timedelta(
                        seconds=((current_begin - yesterday_end).seconds // 2)) - timedelta(
                        hours=UNCERTAIN_WIN_HOURS_HALF)
                    uncertain_end = current_begin - timedelta(
                        seconds=((current_begin - yesterday_end).seconds // 2)) + timedelta(
                        hours=UNCERTAIN_WIN_HOURS_HALF)
                    day_plan.set_valid_begin_datetime(uncertain_end)
                    day_plan.set_uncertain_punch_in_begin_datetime(uncertain_begin)
                    work_day_before.set_valid_end_datetime(uncertain_begin)
                    work_day_before.set_uncertain_punch_out_end_datetime(uncertain_end)
            else:
                day_plan.set_valid_begin_datetime(
                    current_begin - timedelta(hours=NO_PLAN_VALID_EXPAND_HOURS))
                day_plan.set_uncertain_punch_in_begin_datetime(
                    current_begin - timedelta(hours=NO_PLAN_UNCERTAIN_EXPAND_HOURS))
            current_end = day_plan.get_plan_end_datetime()
            day_plan.set_valid_end_datetime(current_end + timedelta(hours=NO_PLAN_VALID_EXPAND_HOURS))
            day_plan.set_uncertain_punch_out_end_datetime(
                current_end + timedelta(hours=NO_PLAN_UNCERTAIN_EXPAND_HOURS))

    def add_punch(self, punch):
        self.punches.append(punch)


class WorkDay(object):
    def __init__(self, work_date, plan_work):
        self.workDate = work_date
        self.planWork = plan_work  # PlanType
        self.validBeginDatetime = get_date_time(work_date, plan_work.get_begin_time())
        if plan_work.is_across_day():
            self.validEndDatetime = get_date_time(work_date + timedelta(days=1),
                                                  plan_work.get_end_time())
        else:
            self.validEndDatetime = get_date_time(work_date, plan_work.get_end_time())
        self.needPunchIn = True
        self.havePunchIn = False
        self.punchInLate = False
        self.punchIn = None
        self.punchInLatest = None
        self.needPunchOut = True
        self.havePunchOut = False
        self.punchOutEarly = False
        self.punchOut = None
        self.uncertainPunchInBeginDatetime = get_date_time(work_date, plan_work.get_begin_time())
        if plan_work.is_across_day():
            self.uncertainPunchOutEndDatetime = get_date_time(work_date + timedelta(days=1),
                                                              plan_work.get_end_time())
        else:
            self.uncertainPunchOutEndDatetime = get_date_time(work_date, plan_work.get_end_time())
        self.uncertainPunchInList = []
        self.uncertainPunchOutList = []
        self.notPunchInRow = None
        self.notPunchOutRow = None

    def punch(self, punch):
        punch_datetime = punch.punchDatetime
        if self.uncertainPunchInBeginDatetime <= punch_datetime <= self.get_plan_begin_datetime():
            self.havePunchIn = True
            self.punchInLate = False
            if not self.punchIn or self.punchIn.punchDatetime > punch_datetime:
                self.punchIn = punch
            self.punchInLatest = punch
            self.clear_uncertain_punch_in()
        elif self.uncertainPunchOutEndDatetime >= punch_datetime >= self.get_plan_end_datetime():
            self.havePunchOut = True
            self.punchOutEarly = False
            if not self.punchOut or self.punchOut.punchDatetime < punch_datetime:
                self.punchOut = punch
            self.clear_uncertain_punch_out()
        else:
            if not self.havePunchIn and (not self.uncertainPunchInList or (
                        (punch_datetime - self.get_plan_begin_datetime()).seconds < (
                                self.get_plan_end_datetime() - punch_datetime).seconds)):
                self.havePunchIn = True
                self.punchInLate = True
                if not self.punchIn or self.punchIn.punchDatetime > punch_datetime:
                    self.punchIn = punch
                self.punchInLatest = punch
            elif (not (self.havePunchOut and not self.punchOutEarly)) and (
                        not self.punchInLatest or not is_same_time_punch(self.punchInLatest, punch)):
                self.havePunchOut = True
                self.punchOutEarly = True
                if not self.punchOut or self.punchOut.punchDatetime < punch_datetime:
                    self.punchOut = punch

    def uncertain_punch_in(self, punch):
        self.uncertainPunchInList.append(punch)

    def uncertain_punch_out(self, punch):
        self.uncertainPunchOutList.append(punch)

    def clear_uncertain_punch_in(self):
        self.uncertainPunchInList = []

    def remove_processed_uncertain_punch_in(self, from_datetime):
        old_punch_in_list = self.uncertainPunchInList
        self.uncertainPunchInList = []
        for punch in old_punch_in_list:
            if can_be_in_out_diff_datetime(from_datetime, punch.punchDatetime):
                self.uncertainPunchInList.append(punch)

    def clear_uncertain_punch_out(self):
        self.uncertainPunchOutList = []

    def get_work_date(self):
        return self.workDate

    def get_plan_type(self):
        return self.planWork

    def have_punch_in(self):
        return self.havePunchIn

    def have_punch_out(self):
        return self.havePunchOut

    def is_punch_in_late(self):
        return self.punchInLate

    def is_punch_out_early(self):
        return self.punchOutEarly

    def get_punch_in_datetime(self):
        if not self.punchIn:
            return None
        else:
            return self.punchIn.punchDatetime

    def get_punch_out_datetime(self):
        if not self.punchOut:
            return None
        else:
            return self.punchOut.punchDatetime

    def is_before_work_uncertain_time(self, punch):
        return punch.punchDatetime < self.uncertainPunchInBeginDatetime

    def is_after_work_uncertain_time(self, punch):
        return punch.punchDatetime > self.uncertainPunchOutEndDatetime

    def is_before_work_valid_time(self, punch):
        return punch.punchDatetime < self.validBeginDatetime

    def is_after_work_valid_time(self, punch):
        return punch.punchDatetime > self.validEndDatetime

    def set_valid_begin_datetime(self, valid_begin_datetime):
        self.validBeginDatetime = valid_begin_datetime

    def set_valid_end_datetime(self, valid_end_datetime):
        self.validEndDatetime = valid_end_datetime

    def set_uncertain_punch_in_begin_datetime(self, uncertain_punch_in_begin_datetime):
        self.uncertainPunchInBeginDatetime = uncertain_punch_in_begin_datetime

    def set_uncertain_punch_out_end_datetime(self, uncertain_punch_out_end_datetime):
        self.uncertainPunchOutEndDatetime = uncertain_punch_out_end_datetime

    def get_plan_begin_datetime(self):
        return get_date_time(self.workDate, self.planWork.get_begin_time())

    def get_plan_end_datetime(self):
        if self.planWork.is_across_day():
            return get_date_time(self.workDate + timedelta(days=1), self.planWork.get_end_time())
        else:
            return get_date_time(self.workDate, self.planWork.get_end_time())


class RestDay(object):
    def __init__(self, work_date, plan):
        self.workDate = work_date
        self.plan = plan  # PlanType
        self.haveOutput = False

    def mark_output(self):
        self.haveOutput = True

    def get_plan_describe(self):
        return self.plan.describe

    def get_plan_begin_datetime(self):
        return get_date_time(self.workDate, self.plan.get_begin_time())

    def get_plan_end_datetime(self):
        return get_date_time(self.workDate, self.plan.get_end_time())


class Punch(object):
    def __init__(self, punch_type, punch_datetime, not_real=False):
        self.notReal = not_real  # 是否是打卡记录中存在的。或者是系统自动添加的True。
        self.punchType = punch_type
        self.punchDatetime = punch_datetime
        self.outputToDetails = False


def is_same_time_punch(punch1, punch2):
    return is_same_time(punch1.punchDatetime, punch2.punchDatetime)


def is_same_time(datetime1, datetime2):
    if datetime1 <= datetime2 < (datetime1 + timedelta(minutes=ONCE_PUNCH_DIFF_MAX_MINUTE)):
        return True
    elif datetime1 >= datetime2 > (datetime1 - timedelta(minutes=ONCE_PUNCH_DIFF_MAX_MINUTE)):
        return True
    else:
        return False


def can_be_in_out_diff_datetime(first_datetime, second_datetime):
    return (second_datetime - first_datetime) > timedelta(hours=PUNCH_TYPE_DIFF_MIN_HOUR)


def can_be_in_out_diff_punch_type(first_punch, second_punch):
    return can_be_in_out_diff_datetime(first_punch.punchDatetime, second_punch.punchDatetime)


def get_date_time(date_obj, time_obj):
    return datetime.combine(date_obj, time_obj)


PunchTypeKey = enum(
    PunchIn='上班签到', PunchOut='下班签退'
)
