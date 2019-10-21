import os
import sys
import time
import csv
import smtplib
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from email import encoders

sender_host = 'smtp.qq.com'  # 默认服务器地址及端口
sender_user = '1206064486@qq.com'
sender_pwd = 'ntidvudghbqgfjhg'
sender_name = u'Jason'


attach_path = r'D:\File\attach'   # 附件所在文件夹
attach_type = ".xlsx"      # 附件后缀名，即类型
addrBook = r'D:\File\receivers.csv'  # 邮件地址通讯录
content_path = r"D:\File\content.txt"   # 邮箱正文内容.txt
month = 201910 # 月份，每次发送前修改


# 根据输入的CSV文件，获取通讯录人名和相应的邮箱地址
# addrs = {name : value}

def getAddrBook(addrBook):
    with open(addrBook, 'r', encoding='gbk') as addrFile:
        reader = csv.reader(addrFile)
        name = []
        value = []
        addrs = {}
        for row in reader:
            name.append(row[0])
            value.append(row[1])
    addrs = dict(zip(name, value))
    return addrs


# 根据附件名称中获得的人名，查找通讯录，找到对应的邮件地址

def getRecvAddr(addrs, person_name):
    if not person_name in addrs:
        print("没有<" + person_name + ">的邮箱地址! 请在联系人中添加此人邮箱后重试。")
        anykey = input("请按任意数字键【0-9】退出程序：")
        if anykey != '':
            time.sleep(1)
            sys.exit(0)
    return addrs[person_name]


# 加载邮件内容

mail_content = 'Test Test'


# 添加附件

def addAttch(attach_file):
    att = MIMEBase('application', 'octet-stream')  # 这两个参数不知道啥意思，二进制流文件
    att.set_payload(open(attach_file, 'rb').read())
    # 此时的附件名称为****.xlsx，截取文件名
    att.add_header('Content-Disposition', 'attachment', filename=(attach_file.split("\\")[-1].split(".")[0] + '-' + str(month) + '.xlsx'))
    encoders.encode_base64(att)
    return att


# 发送邮件
def mailSend():
    smtp = smtplib.SMTP()  # 新建smtp对象
    smtp.connect(sender_host)
    smtp.login(sender_user, sender_pwd)

    for root, dirs, files in os.walk(attach_path):
        for attach_file in files:  # attach_file : ***_2_***.xlsx
            msg = MIMEMultipart('alternative')
            att_name = attach_file.split(".")[0] 
            subject = att_name + '-' + str(month)
            msg['Subject'] = subject  # 设置邮件主题
            person_name = att_name # subject.split("_")[-1]

            addrs = getAddrBook(addrBook)
            recv_addr = getRecvAddr(addrs, person_name)

            msg['From'] = formataddr([sender_name, sender_user])  # 设置发件人名称
            # msg['To'] = person_name # 设置收件人名称
            msg['To'] = formataddr([person_name, recv_addr])  # 设置收件人名称
            
            msg.attach(MIMEText(mail_content,'plain','utf-8'))  # 正文  MIMEText(content,'plain','utf-8')
            attach_file = root + "\\" + attach_file
            att = addAttch(attach_file)
            msg.attach(att)  # 附件
            smtp.sendmail(sender_user, [recv_addr, ], msg.as_string())  # smtp.sendmail(from_addr, to_addrs, msg)
            print("已发送： " + person_name + " <" + recv_addr + ">")
        smtp.quit()

if __name__ == '__main__':
    print("Jason")
    mailSend()
    anykey = input("请按回车键（Enter）退出程序：")
    if anykey:
        exit(50)
