# -*- coding: utf-8 -*-
import csv
import datetime
import os
import xlrd
import codecs
import time, datetime
from random import Random
import uuid


DONE_UnitUID = ""
DONE_CreatedTime = ""
DONE_ALARMUID = ""

weekindex = {
    '星期一':1,
    '星期二':2,
    '星期三':3,
    '星期四':4,
    '星期五':5,
    '星期六':6,
    '星期七':7,
}
classbegin = {
    '第1节':'080000',
    '第2节':'085500',
    '第3节':'100000',
    '第4节':'105500',
    '第5节':'140000',
    '第6节':'145500',
    '第7节':'160000',
    '第8节':'165500',
    '第9节':'190000',
    '第10节':'195500',
    '第11节':'210000',
}

classend = {
    '第1节':'084500',
    '第2节':'094000',
    '第3节':'104500',
    '第4节':'114000',
    '第5节':'144500',
    '第6节':'154000',
    '第7节':'164500',
    '第8节':'174000',
    '第9节':'194500',
    '第10节':'204000',
    '第11节':'213000',
}


def get_week_num (star):
    if ((star) in weekindex):
        week_num = weekindex[star]
    else:
        week_num = 0
    return week_num

def jieci2time_begin (jieci):
    if (jieci in classbegin):
        timestr = classbegin[jieci]

    else:
        pass
    return timestr

def jieci2time_end (jieci):
    if (jieci in classend):
        timestr = classend[jieci]
        # return timestr
    else:
        pass
    return timestr

def Create_T_INFO():

	global DONE_CreatedTime
	date = datetime.datetime.now().strftime("%Y%m%dT%H%M%S")
	DONE_CreatedTime = date + "Z"

	global DONE_ALARMUID
	DONE_ALARMUID = str(uuid.uuid4())

	global DONE_UnitUID
	DONE_UnitUID = str(uuid.uuid5(uuid.NAMESPACE_DNS, "CTGU"))

def getweek_range (astr):
    str1 = astr.split(',')
    str2 = []
    str3 = []
    time = len(str1)
    for i in range(time) :
        str2.append(str1[i].split('周')[0])
    time = len(str2)
    for i in range (time) :
        str3.append(str2[i].split('-'))
    for i in range (time) :
        for j in range(len(str3[i])) :
            str3[i][j] = int (str3[i][j])
    return str3

def getweek (str_get):
    str3 = getweek_range(str_get)
    week = []
    time = len(str3)
    for i in range (time) :
        start = int(str3[i][0])
        if (len(str3[i]) == 2):
            end = int(str3[i][1])
        else:
            end = start
        for j in range(start,end + 1) :
            week.append(j)
    return week

def read_csv (path):
    class_num = 0;
    class_list = []
    with open(path,encoding='GB2312') as csvfile:
        reader = csv.DictReader(csvfile) # 注意函数是大写
        
        for row in reader: # 读取csv文件中的课程表，用class_list存取一个字典数组。
            class_info = {'课程号':'','课程名':'','上课周次':'','上课星期':'','开始节次':'','结束节次':'','上课教师':'','教室名称':''}
            class_info['课程号'] = row['课程号']
            class_info['课程名'] = row['课程名']
            class_info['上课周次'] = getweek(row['上课周次'])
            class_info['周次范围'] = getweek_range(row['上课周次'])
            class_info['上课星期'] = row['上课星期']
            class_info['开始节次'] = row['开始节次']
            class_info['结束节次'] = row['结束节次']
            class_info['上课教师'] = row['上课教师']
            class_info['教室名称'] = row['教室名称']
            class_num = class_num + 1
            class_list.append(class_info)
    class_info = {'class_num':class_num,'class_list':class_list}
    return class_info

def read_xls (path):
    class_num = 0
    class_list = []
    workbook = xlrd.open_workbook(path)

    worksheet = workbook.sheet_by_index(0)

    for i in range(1,worksheet.nrows):
        class_info = {'课程号':'','课程名':'','上课周次':'','上课星期':'','开始节次':'','结束节次':'','上课教师':'','教室名称':''}
        class_info['课程号'] = worksheet.cell_value(i,0)
        class_info['课程名'] = worksheet.cell_value(i,1)
        class_info['上课周次'] = getweek( worksheet.cell_value(i,5) )
        class_info['周次范围'] = getweek_range( worksheet.cell_value(i,5) )
        class_info['上课星期'] = worksheet.cell_value(i,6)
        class_info['开始节次'] = worksheet.cell_value(i,7)
        class_info['结束节次'] = worksheet.cell_value(i,8)
        class_info['上课教师'] = worksheet.cell_value(i,9)
        class_info['教室名称'] = worksheet.cell_value(i,10)
        class_num = class_num + 1
        class_list.append(class_info)
    class_info = {'class_num':class_num,'class_list':class_list}
    return class_info

def checkdate (date_begin,delta_week,number) :
    d1 = datetime.datetime.strptime(date_begin, '%Y-%m-%d')
    delta = datetime.timedelta(days=(delta_week - 1) * 7 + (number - 1))
    d2 = d1 + delta
    date_str = d2.strftime('%Y-%m-%d')
    # print(date_str)
    date_dir = {'year':int(date_str[0:4]),'month':int(date_str[5:7]),'day':int(date_str[8:10])}
    date_str = d2.strftime('%Y%m%d')
    return date_str




def writeisc(date_begin,class_info,wpath):
    if not os.path.exists(wpath):
        os.makedirs(wpath)
    os.chdir(wpath)

    # 删除旧文件
    os.remove("./class.ics")

    head_str = ('BEGIN:VCALENDAR\n'
                'METHOD:PUBLISH\n'
                'VERSION:2.0\n'
                'X-WR-CALNAME:课程表\n'
                'PRODID:-//Apple Inc.//Mac OS X 10.12//EN\n'
                'X-APPLE-CALENDAR-COLOR:#FC4208\n'
                'X-WR-TIMEZONE:Asia/Shanghai\n'
                'CALSCALE:GREGORIAN\n'
                'BEGIN:VTIMEZONE\n'
                'TZID:Asia/Shanghai\n'
                'BEGIN:STANDARD\n'
                'TZOFFSETFROM:+0900\n'
                'RRULE:FREQ=YEARLY;UNTIL=19910914T150000Z;BYMONTH=9;BYDAY=3SU\n'
                'DTSTART:19890917T000000\n'
                'TZNAME:GMT+8\n'
                'TZOFFSETTO:+0800\n'
                'END:STANDARD\n'
                'BEGIN:DAYLIGHT\n'
                'TZOFFSETFROM:+0800\n'
                'DTSTART:19910414T000000\n'
                'TZNAME:GMT+8\n'
                'TZOFFSETTO:+0900\n'
                'RDATE:19910414T000000\n'
                'END:DAYLIGHT\n'
                'END:VTIMEZONE\n\n')
    f = codecs.open("./class.ics", 'a', encoding='utf-8')
    f.writelines (head_str)
    f.close()
    Create_T_INFO()
    global DONE_ALARMUID, DONE_UnitUID, DONE_CreatedTime
    for i in range(class_info['class_num']):

        for j in range(len(class_info['class_list'][i]['周次范围'])) :

            file_name = "./class.ics"
            with codecs.open(file_name,'a', encoding='utf-8') as file:
                if(len(class_info['class_list'][i]['周次范围'][j]) == 1) :
                    count = 1
                else :
                    count = class_info['class_list'][i]['周次范围'][j][1] - class_info['class_list'][i]['周次范围'][j][0] + 1
                
                count_str = str(count)
                begin_date_str = checkdate(date_begin, class_info['class_list'][i]['周次范围'][j][0] , get_week_num(class_info['class_list'][i]['上课星期']) )
                begin_time_str = jieci2time_begin(class_info['class_list'][i]['开始节次'])
                end_time_str = jieci2time_end(class_info['class_list'][i]['结束节次'])

                begin_str = begin_date_str + 'T' + begin_time_str
                end_str = begin_date_str + 'T' + end_time_str


                str_1 = (
                    'BEGIN:VEVENT\n'
                    'CREATED:' + DONE_CreatedTime + '\n'
                )

                str_2 = 'UID:' + str(uuid.uuid4()) + '\n'

                str_3 = 'DTEND;TZID=Asia/Shanghai:' + end_str + '\n' + 'RRULE:FREQ=WEEKLY;INTERVAL=1;COUNT=' + count_str + '\n'
                
                str_4 = 'TRANSP:OPAQUE\nX-APPLE-TRAVEL-ADVISORY-BEHAVIOR:AUTOMATIC\nSUMMARY:' + class_info['class_list'][i]['课程名'] + '\n'

                str_5 = 'LOCATION:' + class_info['class_list'][i]['教室名称'] + '\n'
                str_6 = 'DESCRIPTION:' + class_info['class_list'][i]['上课教师'] + '\n'

                str_7 = 'DTSTART;TZID=Asia/Shanghai:' + begin_str + '\n' + 'DTSTAMP:' + DONE_CreatedTime + '\n'

                str_8 = "SEQUENCE:0\nBEGIN:VALARM\nX-WR-ALARMUID:" + DONE_ALARMUID + '\n' + 'UID:' +  DONE_UnitUID + '\n'

                str_9 = 'TRIGGER:NULL' + '\n' + 'ACTION:DISPLAY\nEND:VALARM\nEND:VEVENT\n\n\n'


                str_final = str_1 + str_2 + str_3 + str_4 + str_5 + str_6 + str_7 + str_8 + str_9          
                file.writelines(str_final)
                file.close()
    f = codecs.open("./class.ics", 'a', encoding='utf-8')
    f.writelines ('END:VCALENDAR\n\n\n')
    f.close()    


# writeisc("2023-02-16", read_xls("./class.xlsx"), "./")
