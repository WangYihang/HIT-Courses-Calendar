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
  4. `python main.py timetable.xls timetable.cvs`
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

