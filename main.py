#!/usr/bin/env python
# encoding:utf-8
# -*- coding: utf-8 -*-

import xlrd
import re
import datetime
import argparse

schedule_normal = [
    ("08:00 AM", "09:45 AM"),  # 1,2
    ("10:00 AM", "11:45 AM"),  # 3,4
    ("01:45 PM", "03:30 PM"),  # 5,6
    ("03:45 PM", "05:30 PM"),  # 7,8
    ("06:30 PM", "08:15 PM"),  # 9,10
    ("08:30 PM", "10:15 PM"),  # 11,12
]

schedule_exam = [
    ("08:00 AM", "10:00 AM"),  # 1,2
    ("10:00 AM", "12:00 PM"),  # 3,4
    ("01:00 PM", "03:00 PM"),  # 5,6
    ("03:45 PM", "05:45 PM"),  # 7,8
    ("06:30 PM", "08:30 PM"),  # 9,10
]

schedule_experiment = [
    ("07:20 AM", "09:50 AM"),  # 1,2
    ("10:00 AM", "12:30 PM"),  # 3,4
    ("01:00 PM", "03:30 PM"),  # 5,6
    ("03:40 PM", "06:10 PM"),  # 7,8
    ("06:30 PM", "09:00 PM"),  # 9,10
]

keys = [
    "Subject",
    "Start Date",
    "Start Time",
    "End Date",
    "End Time",
    "All Day Event",
    "Description",
    "Location",
    "Private",
]


class Subject():
    '''
    Reference: https://support.google.com/calendar/answer/37118?hl=zh-Hans
    '''

    def __init__(self,
                 subject,
                 start_date,
                 start_time,
                 end_date,
                 end_time,
                 all_day_event,
                 description,
                 location,
                 private):
        self.subject = subject
        self.start_date = start_date
        self.start_time = start_time
        self.end_date = end_date
        self.end_time = end_time
        self.all_day_event = all_day_event
        self.description = description
        self.location = location
        self.private = private

    def __str__(self):
        return "%s, %s, %s, %s, %s, %s, %s, %s, %s" % (
            self.subject,
            self.start_date,
            self.start_time,
            self.end_date,
            self.end_time,
            self.all_day_event,
            self.description,
            self.location,
            self.private,
        )


def pretty_print(data):
    if isinstance(data, list):
        print "[",
        for i in data:
            pretty_print(i)
            print ",",
        print "]"
    elif isinstance(data, dict):
        print "{",
        for i in data.items():
            pretty_print(i)
            print ",",
        print "}"
    elif isinstance(data, tuple):
        print "(",
        for i in data:
            pretty_print(i)
            print ":",
        print ")"
    else:
        print data,


def parseExcel(inputfile, outputfile, semester_start_date):
    print "[+] Parsing xls data..."
    with open(outputfile, "w") as f:
        f.write("".join([i + "," for i in keys]) + "\n")
    try:
        work_book = xlrd.open_workbook(inputfile)
    except Exception as e:
        print e
        exit(2)
    sheet_names = work_book.sheet_names()
    for sheet_name in sheet_names:
        sheet = work_book.sheet_by_name(sheet_name)
        nrows = sheet.nrows
        ncols = sheet.ncols
        count = 0
        for r in range(2, nrows):
            for c in range(2, ncols):
                item = sheet.cell_value(r, c)
                lectures_info = parseItem(item.encode("utf-8"))
                for lecture_info in lectures_info:
                    if "[考试]" in lecture_info['name']:
                        lecture_info['start_time'] = schedule_exam[r - 2][0]
                        lecture_info['end_time'] = schedule_exam[r - 2][1]
                    elif "(实验)" in lecture_info['name']:
                        lecture_info['start_time'] = schedule_experiment[r - 2][0]
                        lecture_info['end_time'] = schedule_experiment[r - 2][1]
                    else:
                        lecture_info['start_time'] = schedule_normal[r - 2][0]
                        lecture_info['end_time'] = schedule_normal[r - 2][1]
                    weeks = lecture_info['weeks']
                    # pretty_print(lecture_info)
                    for week in weeks:
                        start_date = semester_start_date + datetime.timedelta(weeks=week - 1) + datetime.timedelta(
                            days=c - 2)
                        end_date = start_date
                        subject = Subject(
                            subject=lecture_info['name'],
                            start_date=start_date,
                            start_time=lecture_info['start_time'],
                            end_date=end_date,
                            end_time=lecture_info['end_time'],
                            all_day_event=False,
                            description=lecture_info['instructor'],
                            location=lecture_info['location'],
                            private=True,
                        )
                        count += 1
                        with open(outputfile, "a+") as f:
                            f.write(str(subject) + "\n")
                # pretty_print(lectures_info)
    print "[+] Output finished!"
    return count


def parseItem(content):
    if len(content) == 0:
        return []
    content_items = re.split("</br>|<br/>", content)
    results = []
    count = 0
    flag = 0
    distribute = []
    clainfo = []
    while count < len(content_items):
        if "周" in content_items[count]:
            if flag == 0:
                clainfo.append(content_items[count - 1])
                flag = 1
            clainfo.append(content_items[count])
            if content_items[count][len(content_items[count]) - 1].isdigit():
                distribute.append(clainfo)
                clainfo = []
                flag = 0
        if content_items[count][len(content_items[count]) - 1].isdigit():
            if flag == 1:
                clainfo.append(content_items[count])
                distribute.append(clainfo)
                clainfo = []
                flag = 0
        count = count + 1
    for k in range(0, len(distribute), 1):
        name = distribute[k][0]
        if "[考试]" not in name:
            if len(distribute[k]) == 2:
                left_data = distribute[k][1]
                instructor = left_data.split("[")[0]

                # regex
                # left_data = "学术英语写作(提高)</br>姚静[1-3，5-16]周正心432，[4]周校外"
                # weeks_info_temp = re.compile(r'(?<=\[).*(?<!\])').findall(left_data)
                # print left_data
                weeks_info_temp = re.compile(r'(?<=\[).*?(?=\])').findall(left_data)
                # print weeks_info_temp
                weeks_info = []
                for i in weeks_info_temp:
                    weeks_info += i.split("，")

                # weeks_info = left_data.split("[")[1].split("]")[0].split('，')

                weeks = []

                for week_info in weeks_info:
                    if "-" in week_info:
                        start_week = int(week_info.split("-")[0])
                        end_week = int(week_info.split("-")[1])
                    else:
                        start_week = int(week_info)
                        end_week = start_week

                    if "双]周" in left_data:
                        location = left_data.split("双]周")[1].split("，")[0]
                        for i in range(start_week, end_week):
                            if i % 2 == 0:
                                weeks.append(i)
                    elif "]单周" in left_data:
                        location = left_data.split("]单周")[1].split("，")[0]
                        for i in range(start_week, end_week):
                            if i % 2 == 1:
                                weeks.append(i)
                    elif "]周" in left_data:
                        # .split("，")[0] for 正心432，[4]周校外
                        location = left_data.split("]周")[1].split("，")[0]
                        for i in range(start_week, end_week + 1):
                            weeks.append(i)
                    else:
                        exit(3)

                lecture_info = {
                    "name": name,
                    # "is_examination": is_examination,
                    "instructor": instructor,
                    "location": location,
                    "weeks": weeks,
                }
                results.append(lecture_info)
            if len(distribute[k]) == 3:
                location = distribute[k][2]
                instrs = distribute[k][1].split("周，")
                for l in range(0, len(instrs), 1):
                    instructor = instrs[l].split("[")[0]
                    weeks_info_temp = instrs[l].split("[")[1]
                    weeks_info_temp = weeks_info_temp.split("]")[0]
                    weeks_info_temp = weeks_info_temp.split("双")[0]
                    weeks_info = []
                    if "，" in weeks_info_temp:
                        weeks_info += weeks_info_temp.split("，")
                    if "，" not in weeks_info_temp:
                        weeks_info.append(weeks_info_temp)
                    weeks = []
                    for week_info in weeks_info:
                        if "-" in week_info:
                            start_week = int(week_info.split("-")[0])
                            end_week = int(week_info.split("-")[1])
                        else:
                            start_week = int(week_info)
                            end_week = start_week
                        if "双]周" in instrs[l]:

                            for i in range(start_week, end_week):
                                if i % 2 == 0:
                                    weeks.append(i)
                        elif "双]" in instrs[l]:

                            for i in range(start_week, end_week):
                                if i % 2 == 0:
                                    weeks.append(i)
                        elif "]单周" in instrs[l]:

                            for i in range(start_week, end_week):
                                if i % 2 == 1:
                                    weeks.append(i)
                        elif "]单" in instrs[l]:

                            for i in range(start_week, end_week):
                                if i % 2 == 1:
                                    weeks.append(i)
                        elif "]周" in instrs[l]:
                            # .split("，")[0] for 正心432，[4]周校外

                            for i in range(start_week, end_week + 1):
                                weeks.append(i)
                        elif "]" in instrs[l]:
                            # .split("，")[0] for 正心432，[4]周校外

                            for i in range(start_week, end_week + 1):
                                weeks.append(i)
                        else:
                            exit(3)

                    lecture_info = {
                        "name": name,
                        # "is_examination": is_examination,
                        "instructor": instructor,
                        "location": location,
                        "weeks": weeks,
                    }
                    results.append(lecture_info)
            if len(distribute[k]) > 3:
                infos = ""
                for l in range(1, len(distribute[k]), 1):
                    infos += distribute[k][l]
                informa = infos.split("，")
                for k in range(0, len(informa), 1):
                    instructor = informa[k].split("[")[0]
                    loc = informa[k].split("周")
                    location = loc[len(loc) - 1]
                    weeks_info_temp = informa[k].split("[")[1]
                    weeks_info_temp = weeks_info_temp.split("]")[0]
                    weeks_info_temp = weeks_info_temp.split("双")[0]
                    weeks_info = []
                    if "，" in weeks_info_temp:
                        weeks_info += weeks_info_temp.split("，")
                    if "，" not in weeks_info_temp:
                        weeks_info.append(weeks_info_temp)
                    weeks = []
                    for week_info in weeks_info:
                        if "-" in week_info:
                            start_week = int(week_info.split("-")[0])
                            end_week = int(week_info.split("-")[1])
                        else:
                            start_week = int(week_info)
                            end_week = start_week
                        if "双]周" in informa[k]:

                            for i in range(start_week, end_week):
                                if i % 2 == 0:
                                    weeks.append(i)
                        elif "双]" in informa[k]:

                            for i in range(start_week, end_week):
                                if i % 2 == 0:
                                    weeks.append(i)
                        elif "]单周" in informa[k]:

                            for i in range(start_week, end_week):
                                if i % 2 == 1:
                                    weeks.append(i)
                        elif "]单" in informa[k]:

                            for i in range(start_week, end_week):
                                if i % 2 == 1:
                                    weeks.append(i)
                        if "]周" in informa[k]:
                            # .split("，")[0] for 正心432，[4]周校外

                            for i in range(start_week, end_week + 1):
                                weeks.append(i)
                        elif "]" in informa[k]:
                            # .split("，")[0] for 正心432，[4]周校外

                            for i in range(start_week, end_week + 1):
                                weeks.append(i)
                        else:
                            exit(3)

                    lecture_info = {
                        "name": name,
                        # "is_examination": is_examination,
                        "instructor": instructor,
                        "location": location,
                        "weeks": weeks,
                    }
                    results.append(lecture_info)
        if "[考试]" in name:
            infos = distribute[k][1].split(" ")
            location = ""
            instructor = ""
            weeks_info = ""
            weeks_info = infos[0]
            weeks_info = weeks_info.split("[")[1]
            weeks_info = weeks_info.split("]")[0]
            weeks = []
            weeks = [int(weeks_info)]
            if len(infos) == 2:
                location = infos[1]
            lecture_info = {
                "name": name,
                # "is_examination": is_examination,
                "instructor": instructor,
                "location": location,
                "weeks": weeks,
            }
            results.append(lecture_info)
    return results


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-i",
        "--inputfile",
        help="input file to convert",
        default="timetable6.xls"
    )
    parser.add_argument(
        "-o",
        "--outputfile",
        help="output file to save",
        default="timetable.cvs"
    )
    parser.add_argument(
        "-s",
        "--semester",
        help="semester start date, format: year/month/day, example: 2018/02/26",
        default="2018/09/03"
    )
    args = parser.parse_args()

    # BUG [1-14]双周
    year = int(args.semester.split("/")[0])
    month = int(args.semester.split("/")[1])
    day = int(args.semester.split("/")[2])
    try:
        semester_start_date = datetime.date(year, month, day)
        print "[+] Semester start date: %s" % (semester_start_date)
    except Exception as e:
        print str(e)
        exit(4)
    count = parseExcel(args.inputfile, args.outputfile, semester_start_date)
    print "[+] %d lectures found" % (count)
    print "[+] Please check %s" % (args.outputfile)


if __name__ == '__main__':
    main()

