# -*- coding:utf-8 -*-


import xlrd
import traceback
import re
from datetime import datetime
import time

from qcloudsms_py import SmsSingleSender
from qcloudsms_py.httpclient import HTTPError

from smtplib import SMTP

# 短信应用配置
appid = 1400206808
appkey = "9d727cf12efc14e12ed13a07fdc552a5"
template_id = 325573
sms_sign = ""

# 默认文件名
DEFAULTFILENAME = 'schedule.xls'
LOGFILE = 'log.txt'
try:
    logfp = open(LOGFILE, 'a')
except Exception as e:
    print(e)


def notifyme(subject, msg):
    # print('msg=',msg)
    HOST = "202.200.112.36"  # 定义smtp主机
    SUBJECT = subject #"Jiankao report"  # 定义邮件主题
    TO = "kedong@xaut.edu.cn"  # 定义邮件收件人
    FROM = "kedong@xaut.edu.cn"  # 定义邮件发件人
    text = msg  # 邮件内容,编码为ASCII范围内的字符或字节字符串，所以不能写中文
    BODY = '\r\n'.join((  # 组合sendmail方法的邮件主体内容，各段以"\r\n"进行分离
        "From: %s" % "JiankaoBot",
        "TO: %s" % TO,
        "subject: %s" % SUBJECT,
        "",
        text
    ))
    server = SMTP()  # 创建一个smtp对象
    server.connect(HOST, '25')  # 链接smtp主机
    print('SMTP connected')
    server.login(FROM, "xautmima2015")  # 邮箱账号登陆
    server.sendmail(FROM, TO, BODY)  # 发送邮件
    server.quit()  # 端口smtp链接


# 调用短信息接口给所有监考人发送短信息
def sendAlert(alertList):
    cnt = 0
    okcnt = 0
    logfp.write('SMS List:\n')
    for a in alertList:
        number = a['number']
        params = a['params']
        rst = send_sms(number, params)
        if rst['result'] == 0 and rst['errmsg'] =='OK':
            logfp.write('Number: %s, Message: %s老师，监考小助手提醒您%s的%s在%s有监考任务(%s)，请核实。\n' % (number, params['name'],params['date'],params['time'],params['location'],params['type']))
            okcnt = okcnt +1
        else:
            logfp.write('ERORR: %s' % str(rst))
        cnt = cnt + 1
        time.sleep(0.3)
    return cnt, okcnt


def sendAlertStub(alertList):
    cnt = 0
    okcnt = 0
    # logfp.write('SMS List:\n')
    # for a in alertList:
    #     number = a['number']
    #     params = a['params']
    #     rst = send_sms(number, params)
    #     if rst['result'] == 0 and rst['errmsg'] =='OK':
    #         logfp.write('Number: %s, Message: %s老师，监考小助手提醒您%s的%s在%s有监考任务(%s)，请核实。\n' % (number, params['name'],params['date'],params['time'],params['location'],params['type']))
    #         okcnt = okcnt +1
    #     else:
    #         logfp.write('ERORR: %s' % str(rst))
    #     cnt = cnt + 1
    #     time.sleep(0.3)
    return cnt, okcnt

# 发短信接口调用
def send_sms(number, params):
    ssender = SmsSingleSender(appid, appkey)
    params = [params['name'],params['date'],params['time'],params['location'],params['type']]  # 当模板没有参数时，`params = []`，数组具体的元素个数和模板中变量个数必须一致，例如示例中 templateId:5678 对应一个变量，参数数组中元素个数也必须是一个
    try:
        result = ssender.send_with_param(86, number,
                                         template_id, params, sign=sms_sign, extend="",
                                         ext="")  # 签名参数未提供或者为空时，会使用默认签名发送短信
    except HTTPError as e:
        print(e)
        logfp.write(str(e) + '\n')
    except Exception as e:
        print(e)
        logfp.write(str(e) + '\n')

    return result


# 从文本中解析出监考日期，有两种格式：一种是X周周Y（2019-01-23xxx）；另一种是2019年8月3日
def parseSchedule(schedstr):
    schedule=None
    pattern = ['.*(\d{4})\s*-\s*(\d{1,2})\s*-\s*(\d{1,2}).*', '.*(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日.*']
    for p in pattern:
        obj = re.match(p, schedstr)
        if obj is not None:
            year = int(obj.group(1))
            month = int(obj.group(2))
            day = int(obj.group(3))
            schedule = datetime(year, month, day)
            break
    return schedule


# 从文本中解析出监考时间
def parseTime(schedstr):
    print(schedstr)
    schedtime = None
    pattern=['.*[\(,\s](\d{1,2}\s*:\s*\d{1,2}\s*-\s*\d{1,2}\s*:\s*\d{1,2}).*']
    for p in pattern:
        obj = re.match(p, schedstr)
        if obj is not None:
            schedtime = obj.group(1)
            break

    return schedtime


# 从文件中载入监考信息
def loadData(file=None):
    filename = DEFAULTFILENAME if file is None else file
    try:
        # open xls file
        data = xlrd.open_workbook(filename)
        # get 1st datasheet
        table = data.sheets()[0]

        alldata = []

        nrows = table.nrows
        ncols = table.ncols

        if nrows <= 1 or ncols < 1:
            print('Invalid data.')
            return None

        print('rows:', nrows)
        print('cols:', ncols)

        # skip first row for title
        for r in range(1, nrows):
            row = table.row_values(r)
            # print(row)
            tmp = {'lecturer': row[0].strip(),
                   'course': row[1].strip(),
                   'schedule': parseSchedule(row[2].strip()) if row[2] != '' else '',
                   'time': parseTime(row[2].strip()) if row[2] != '' else '',
                   'location': row[3].strip(),
                   'stunumber': int(row[4]) if row[4] != '' else '',
                   'teacher1': row[5].strip(),
                   'teacher2': row[6].strip(),
                   'teacher3': row[7].strip(),
                   'teacher4': row[8].strip()
                   }
            alldata.append(tmp)

        [print(a) for a in alldata]
        return alldata
    except Exception as e:
        traceback.print_exc()

# 从文件中载入监考人联系方式
def loadContacts(file=None):
    filename = DEFAULTFILENAME if file is None else file
    try:
        # open xls file
        data = xlrd.open_workbook(filename)
        # get 1st datasheet
        table = data.sheets()[1]

        allcontacts={}
        nrows = table.nrows
        ncols = table.ncols

        if nrows <= 1 or ncols < 3:
            print('Invalid data.')
            return None

        for r in range(1,nrows):
            row = table.row_values(r)
            if row[2].strip() == 'Y' or row[2].strip() == 'y' or row[2].strip() == '是':
                # allcontacts[row[0]]=str(int(row[1]))
                allcontacts[row[0]] = {'number': str(int(row[1])), 'nickname': row[3]}

        print(allcontacts)
        return allcontacts

    except Exception as e:
        traceback.print_exc()


# 根据当前的日期判断是否需要提醒
def isAlert(sched):
    # 根据当前的日期和sched的日期判断这条监考场次是否需要通知监考人
    # 规则是，监考当日前三天，一天各通知一次
    if sched is None or 'schedule' not in sched.keys():
        return False

    schedule = sched['schedule']
    today = datetime.today()

    # print((schedule-today).days+1)

    if (schedule-today).days+1 in [1]: # only alert in previous day 20190618
        return True

    return False

# 根据监考信息和联系人生成提醒的列表
def genAlertList(data, contacts):
    alertList=[]

    for sched in data:
        if isAlert(sched):
            if sched['teacher1'] !='' and sched['teacher1'] in contacts.keys():
                alertList.append({
                    'number': contacts[sched['teacher1']]['number'],
                    'nickname': contacts[sched['teacher1']]['nickname'],
                    'params': {'name': sched['teacher1'], 'date': sched['schedule'].strftime('%Y-%m-%d'), 'time': sched['time'], 'location': sched['location'], 'type': '主监'}
                })
            if sched['teacher2'] != '' and sched['teacher2'] in contacts.keys():
                alertList.append({
                    'number': contacts[sched['teacher2']]['number'],
                    'nickname': contacts[sched['teacher2']]['nickname'],
                    'params': {'name': sched['teacher2'], 'date': sched['schedule'].strftime('%Y-%m-%d'),
                               'time': sched['time'], 'location': sched['location'], 'type': '辅监'}
                })
            if sched['teacher3'] != '' and sched['teacher3'] in contacts.keys():
                alertList.append({
                    'number': contacts[sched['teacher3']]['number'],
                    'nickname': contacts[sched['teacher3']]['nickname'],
                    'params': {'name': sched['teacher3'], 'date': sched['schedule'].strftime('%Y-%m-%d'),
                               'time': sched['time'], 'location': sched['location'], 'type': '辅监'}
                })
            if sched['teacher4'] != '' and sched['teacher4'] in contacts.keys():
                alertList.append({
                    'number': contacts[sched['teacher4']]['number'],
                    'nickname': contacts[sched['teacher4']]['nickname'],
                    'params': {'name': sched['teacher4'], 'date': sched['schedule'].strftime('%Y-%m-%d'),
                               'time': sched['time'], 'location': sched['location'], 'type': '辅监'}
                })

    print(alertList)
    logfp.write('Alert List:\n')
    [logfp.write(str(a)) for a in alertList]
    logfp.write('\n')
    return alertList


if __name__ == '__main__':

    logfp.write('\n-------START at %s-------\n' % (datetime.now().strftime('%Y-%m-%d %H:%M:%S')))

    alldata = loadData()
    logfp.write('Alldata:\n')
    [logfp.write(str(d) + '\n') for d in alldata]

    allcontacts = loadContacts()
    logfp.write('AllContacts:\n')
    [logfp.write(str(c) + '\n') for c in allcontacts]

    alertList = genAlertList(alldata, allcontacts)
    # cnt, okcnt = sendAlertStub(alertList)
    cnt, okcnt = sendAlert(alertList)
    logfp.write('%d SMS sent in total %d' % (okcnt, cnt))
    logfp.write('\n-------END-------\n')

    notifyme('JKBot Report %s' % datetime.strftime(datetime.now(), '%Y-%m-%d'),
             '%d SMS sent in total %d at %s\n\nAlertList:\n%s' % (okcnt,
                                                                  cnt,
                                                                  datetime.strftime(datetime.now(), '%Y-%m-%d %H:%M:%S'),
                                                                  ascii(str(alertList)).replace('}, {','\n')))

    logfp.close()




