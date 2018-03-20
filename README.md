哈尔滨工业大学教务处课表 Excel 转换 cvs 脚本
---

#### Installation
```
git clone https://github.com/WangYihang/ClassScheduleOfHIT.git
cd ClassScheduleOfHIT
pip install -r requirements.txt
```

#### Usage:

1. 登录教务处网站
2. 查看课表，并导出课表为 excel
3. 根据使用说明执行脚本  
  1. -i 为 excel 文件所在路径  
  2. -o 为输出 cvs 文件路径  
  3. -s 为本学期开始日期，格式为：年/月/日 （例如：2018/02/26）  
  4. `python main.py -i timetable.xls -o timetable.cvs -s 2018/02/26`    
4. 执行成功后生成 cvs 文件，即可导入 Google Calendar 等日历管理工具

```
usage: main.py [-h] [-i INPUTFILE] [-o OUTPUTFILE] [-s SEMESTER]

optional arguments:
  -h, --help            show this help message and exit
  -i INPUTFILE, --inputfile INPUTFILE
                        input file to convert
  -o OUTPUTFILE, --outputfile OUTPUTFILE
                        output file to save
  -s SEMESTER, --semester SEMESTER
                        semester start date, format: year/month/day, example:
                        2018/02/26
```

#### Example:
timetable.cvs
```
Subject,Start Date,Start Time,End Date,End Time,All Day Event,Description,Location,Private,
软件构造, 2018-02-26, 08:00 AM, 2018-02-26, 09:45 AM, False, 徐汉川, 正心23, True
软件构造, 2018-03-05, 08:00 AM, 2018-03-05, 09:45 AM, False, 徐汉川, 正心23, True
软件构造, 2018-03-12, 08:00 AM, 2018-03-12, 09:45 AM, False, 徐汉川, 正心23, True
软件构造, 2018-03-19, 08:00 AM, 2018-03-19, 09:45 AM, False, 徐汉川, 正心23, True
软件构造, 2018-03-26, 08:00 AM, 2018-03-26, 09:45 AM, False, 徐汉川, 正心23, True
软件构造, 2018-04-02, 08:00 AM, 2018-04-02, 09:45 AM, False, 徐汉川, 正心23, True
软件构造, 2018-04-09, 08:00 AM, 2018-04-09, 09:45 AM, False, 徐汉川, 正心23, True
软件构造, 2018-04-16, 08:00 AM, 2018-04-16, 09:45 AM, False, 徐汉川, 正心23, True
软件构造, 2018-04-23, 08:00 AM, 2018-04-23, 09:45 AM, False, 徐汉川, 正心23, True
...
马克思主义基本原理, 2018-05-25, 01:45 PM, 2018-05-25, 03:30 PM, False, 彭华, 正心32, True
马克思主义基本原理, 2018-06-01, 01:45 PM, 2018-06-01, 03:30 PM, False, 彭华, 正心32, True
运筹学, 2018-02-27, 03:45 PM, 2018-02-27, 05:30 PM, False, 刘绍辉, 正心12, True
运筹学, 2018-03-06, 03:45 PM, 2018-03-06, 05:30 PM, False, 刘绍辉, 正心12, True
运筹学, 2018-03-13, 03:45 PM, 2018-03-13, 05:30 PM, False, 刘绍辉, 正心12, True
运筹学, 2018-03-20, 03:45 PM, 2018-03-20, 05:30 PM, False, 刘绍辉, 正心12, True
运筹学, 2018-03-27, 03:45 PM, 2018-03-27, 05:30 PM, False, 刘绍辉, 正心12, True
运筹学, 2018-04-03, 03:45 PM, 2018-04-03, 05:30 PM, False, 刘绍辉, 正心12, True
运筹学, 2018-04-10, 03:45 PM, 2018-04-10, 05:30 PM, False, 刘绍辉, 正心12, True
运筹学, 2018-04-17, 03:45 PM, 2018-04-17, 05:30 PM, False, 刘绍辉, 正心12, True
```
