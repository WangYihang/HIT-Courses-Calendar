#!/usr/bin/env python
# encoding:utf-8
# -*- coding: utf-8 -*-

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
        print("[")
        for i in data:
            pretty_print(i)
            print(",")
        print("]")
    elif isinstance(data, dict):
        print("{")
        for i in data.items():
            pretty_print(i)
            print(",")
        print("}")
    elif isinstance(data, tuple):
        print("(")
        for i in data:
            pretty_print(i)
            print(":")
        print(")")
    else:
        print(data)


def parseTXT(inputfile, outputfile, semester_start_date):
    subjects = []
    print("[+] Parsing txt data...")
    # The output of wechat tastes like shit.
    # 12周周三12节正心31
    # What the hell? Tell me who is able to recognize the `12` is twelve or `1, 2`?
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
        subjects.append(subject)
        count += 1
    print("[+] Output finished!")
    return subjects

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-i",
        "--inputfile",
        help="input file to convert",
        default="timetable6.txt"
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
        print("[+] Semester start date: %s" % (semester_start_date))
    except Exception as e:
        print(str(e))
        exit(4)
    subjects = parseTXT(args.inputfile, args.outputfile, semester_start_date)
    print("[+] %d lectures found" % (len(subjects)))
    print("[+] Please check %s" % (args.outputfile))


if __name__ == '__main__':
    main()

