#! /usr/bin/env python

import xlrd
import xlwt
import re
import datetime
import json


with open("config.txt", 'r') as f:
    params = json.load(f)
FILE = params["FILE"]
MONTH = params['MONTH']
ON_WORK_TIME = params['ON_WORK_TIME']
LUNCH_TIME = params['LUNCH_TIME']
REST_TIME = params['REST_TIME']
AFTERNOON_WORK_TIME = params['AFTERNOON_WORK_TIME']
OFF_WORK_TIME = params['OFF_WORK_TIME']
OVER_WORK_TIME = params['OVER_WORK_TIME']
OVER_TIME = params['OVER_TIME']

def str_to_absmin(t_str):
    a = list(map(int, t_str.split(':')))  # list() for python3 compatible
    return a[0]*60 + a[1]

def duration(start, end):
    return str_to_absmin(end) - str_to_absmin(start)

def proc_time(time_list, is_weekend=False):
    if len(time_list) == 0:
        return "", "~", 0, 0
    if len(time_list) == 1:
        return "", time_list[0]+"~", 0, 0

    start = time_list[0]
    end = time_list[-1]
    start_min = str_to_absmin(start)
    end_min = str_to_absmin(end)

    tag = ""
    start_end = start + "~" + end
    work_duration = 0
    over_duration = 0

    if is_weekend:
        over_duration = duration(start, end)
        over_duration = round(over_duration/60.0, 1) # * 2)/2.0

        return tag, start_end, work_duration, over_duration

    else:
        morning_work_min = duration(ON_WORK_TIME, LUNCH_TIME)
        afternoon_work_min = duration(AFTERNOON_WORK_TIME, OFF_WORK_TIME)
        regular_work_min =  morning_work_min + afternoon_work_min

        if start_min <= str_to_absmin(ON_WORK_TIME):  # check in regular
            if end_min > str_to_absmin(OVER_TIME):  # work over time
                work_duration = regular_work_min
                over_duration = duration(OVER_WORK_TIME, end)
            elif end_min >= str_to_absmin(OFF_WORK_TIME): # regular work
                work_duration = regular_work_min
            elif end_min >= str_to_absmin(AFTERNOON_WORK_TIME):  # work over lunch
                work_duration = morning_work_min + duration(AFTERNOON_WORK_TIME, end)
            elif end_min >= str_to_absmin(LUNCH_TIME): # work whole morning
                work_duration = morning_work_min
            else: # work only morning
                work_duration = duration(ON_WORK_TIME, end)

        elif start_min > str_to_absmin(ON_WORK_TIME) and start_min <= str_to_absmin(LUNCH_TIME): # late check in morning
            late = start_min - str_to_absmin(ON_WORK_TIME)
            tag = "late: " + str(late) + "min"
            if late < 30: # late but worktime is full
                late = 0
                start = ON_WORK_TIME

            if  late > 60:
                tag = "absence: " + str(late) + "min"

            if end_min > str_to_absmin(OVER_TIME):  # work over time
                work_duration = regular_work_min - late
                over_duration = duration(OVER_WORK_TIME, end)
            elif end_min > str_to_absmin(OFF_WORK_TIME): # regular work
                work_duration = regular_work_min - late
            elif end_min > str_to_absmin(AFTERNOON_WORK_TIME):  # work over lunch
                work_duration = duration(start, LUNCH_TIME) + duration(AFTERNOON_WORK_TIME, end)
            elif end_min >= str_to_absmin(LUNCH_TIME): # work whole morning
                work_duration = duration(start, LUNCH_TIME)
            else: # work only morning
                work_duration = duration(start, end)

        # check in lunchtime
        elif start_min > str_to_absmin(LUNCH_TIME) and start_min < str_to_absmin(AFTERNOON_WORK_TIME):
            tag = "absence: " + str(morning_work_min) + "min"

            if end_min > str_to_absmin(OVER_TIME):  # work over time
                work_duration = afternoon_work_min
                over_duration = duration(OVER_WORK_TIME, end)
            elif end_min > str_to_absmin(OFF_WORK_TIME): # regular work
                work_duration = afternoon_work_min
            elif end_min > str_to_absmin(AFTERNOON_WORK_TIME):  # work over lunch
                work_duration = duration(start, end)
            else:
                pass

        # check in afternoon
        elif start_min > str_to_absmin(AFTERNOON_WORK_TIME) and start_min <= str_to_absmin(OFF_WORK_TIME): # check in afternoon
            tag = "absence: morning"
            if end_min > str_to_absmin(OVER_TIME):  # work over time
                work_duration = duration(start, OFF_WORK_TIME)
                over_duration = duration(OVER_WORK_TIME, end)
            elif end_min > str_to_absmin(OFF_WORK_TIME): # regular work
                work_duration = duration(start, OFF_WORK_TIME)
            else:
                work_duration = duration(start, end)

        else: # check in evening
            if end_min > str_to_absmin(OVER_TIME):  # work over time
                over_duration = duration(OVER_WORK_TIME, end)
            else:
                pass

        work_duration = round(work_duration/60.0, 1) # * 2)/2.0
        over_duration = round(over_duration/60.0, 1) # * 2)/2.0

        return tag, start_end, work_duration, over_duration

def check_weekend(day):
    weekenum = ["Mon", "Tus", "Wed", "Thu", "Fri", "Sat", "Sun"]
    year_month = MONTH.split('/')
    d = datetime.date(int(year_month[0]), int(year_month[1]), int(day))
    if d.weekday() == 5 or d.weekday() == 6:
        return True, weekenum[d.weekday()]
    else:
        return False, weekenum[d.weekday()]

if __name__ == "__main__":
    src_book = xlrd.open_workbook(FILE)
    src_sheet = src_book.sheets()[2]
    n_rows = src_sheet.nrows
    print("sheet rows:{}".format(n_rows))

    dst_book = xlwt.Workbook()
    dst_sheet = dst_book.add_sheet('Sheet1')

    # copy the head
    row = src_sheet.row_values(2)
    dst_sheet.write(0, 0, row[0])
    dst_sheet.write(0, 1, row[2])
    dst_sheet.write(0, 20, "generate by whj")
    row = src_sheet.row_values(3)
    for i, r in enumerate(row):
        dst_sheet.write(1, i+1, r)

    # copy and calc work time
    ind = 2
    for i in range(4, n_rows):
        row = src_sheet.row_values(i)
        if i%2 == 0:
            dst_sheet.write(ind, 0, row[2] + ":".encode('utf-8') + row[10])
            ind += 1
        else:
            # write title
            dst_sheet.write(ind, 0, "start~end")
            dst_sheet.write(ind+1, 0, "worktime")
            dst_sheet.write(ind+2, 0, "overtime")
            dst_sheet.write(ind+3, 0, "comment")
            for j, r in enumerate(row):
                time_list = re.findall(r'.{5}', r)
                is_weekend, day_tag = check_weekend(src_sheet.cell_value(3, j))
                tag, start_end, work_duration, over_duration = proc_time(time_list, is_weekend)
                dst_sheet.write(ind, j+1, start_end)
                dst_sheet.write(ind+1, j+1, work_duration)
                dst_sheet.write(ind+2, j+1, over_duration)
                dst_sheet.write(ind+3, j+1, tag)
                if is_weekend:
                    dst_sheet.write(ind-1, j+1, day_tag)
            ind += 4

    dst_book.save("new.xls")
