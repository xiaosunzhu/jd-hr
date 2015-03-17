# coding=utf-8

from datetime import datetime, timedelta


__author__ = 'yijun.sun'

FLOAT_TYPE = type(1.0)
STRING_TYPE = type('')


def enum(**enums):
    return type('Enum', (), enums)


class PlanType(object):
    def __init__(self, name, begin_time=None, end_time=None, across_day=False):
        self.name = name
        self.acrossDay = across_day
        self.beginTime = begin_time
        self.endTime = end_time
        if not begin_time:
            self.needWork = False
        else:
            self.needWork = True

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
NO_PLAN_EXPAND_HOURS = 12
PUNCH_TYPE_DIFF_MIN_HOUR = 4
ONCE_PUNCH_DIFF_MAX_MINUTE = 10


class Person(object):
    def __init__(self, name, department):
        self.name = name
        self.department = department
        self.workDays = {}  # Map(date,WorkDay)
        self.punches = []  # Map(date,Punch[])

    def add_work_day(self, work_day):
        self.workDays[work_day.workDate] = work_day
        yesterday = work_day.workDate - timedelta(1)

        work_day_before = self.workDays.get(yesterday)
        current_begin = work_day.get_plan_begin_datetime()
        if work_day_before:
            yesterday_end = work_day_before.get_plan_end_datetime()
            uncertain_begin = current_begin - timedelta(
                seconds=((current_begin - yesterday_end).seconds // 2)) - timedelta(
                hours=UNCERTAIN_WIN_HOURS_HALF)
            uncertain_end = current_begin - timedelta(
                seconds=((current_begin - yesterday_end).seconds // 2)) + timedelta(
                hours=UNCERTAIN_WIN_HOURS_HALF)
            work_day.set_valid_begin_datetime(uncertain_end)
            work_day.set_uncertain_punch_in_begin_datetime(uncertain_begin)
            work_day_before.set_valid_end_datetime(uncertain_begin)
            work_day_before.set_uncertain_punch_out_end_datetime(uncertain_end)
        else:
            work_day.set_valid_begin_datetime(current_begin - timedelta(hours=NO_PLAN_EXPAND_HOURS))
            work_day.set_uncertain_punch_in_begin_datetime(
                current_begin - timedelta(hours=NO_PLAN_EXPAND_HOURS))
        current_end = work_day.get_plan_end_datetime()
        work_day.set_valid_end_datetime(current_end + timedelta(hours=NO_PLAN_EXPAND_HOURS))
        work_day.set_uncertain_punch_out_end_datetime(
            current_end + timedelta(hours=NO_PLAN_EXPAND_HOURS))

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
        self.havePunchIn = False
        self.punchInLate = False
        self.punchIn = None
        self.punchInLatest = None
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

    def punch(self, punch):
        punch_datetime = punch.punchDatetime
        punch_type = punch.punchType
        if self.uncertainPunchInBeginDatetime < punch_datetime <= self.get_plan_begin_datetime():
            self.havePunchIn = True
            self.punchInLate = False
            if not self.punchIn or self.punchIn.punchDatetime > punch_datetime:
                self.punchIn = punch
            self.punchInLatest = punch
            self.clear_uncertain_punch_in()
        elif self.uncertainPunchOutEndDatetime > punch_datetime >= self.get_plan_end_datetime():
            self.havePunchOut = True
            self.punchOutEarly = False
            if not self.punchOut or self.punchOut.punchDatetime < punch_datetime:
                self.punchOut = punch
            self.clear_uncertain_punch_out()
        else:
            if not self.havePunchIn:
                self.havePunchIn = True
                self.punchInLate = True
                if not self.punchIn or self.punchIn.punchDatetime > punch_datetime:
                    self.punchIn = punch
                self.punchInLatest = punch
            elif not is_same_time_punch(self.punchInLatest, punch):
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
            if (from_datetime + timedelta(hours=PUNCH_TYPE_DIFF_MIN_HOUR)) < punch.punchDatetime:
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


class Punch(object):
    def __init__(self, punch_type, punch_datetime):
        self.punchType = punch_type
        self.punchDatetime = punch_datetime
        self.outputToDetails = False


def is_same_time_punch(punch1, punch2):
    if punch1.punchDatetime <= punch2.punchDatetime < (
            punch1.punchDatetime + timedelta(minutes=ONCE_PUNCH_DIFF_MAX_MINUTE)):
        return True
    elif punch1.punchDatetime >= punch2.punchDatetime > (
            punch1.punchDatetime - timedelta(minutes=ONCE_PUNCH_DIFF_MAX_MINUTE)):
        return True


def get_date_time(date_obj, time_obj):
    return datetime.combine(date_obj, time_obj)


PunchTypeKey = enum(
    PunchIn='上班签到', PunchOut='下班签退'
)