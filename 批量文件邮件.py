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
sender_name = u'Jason Huang'

attach_path = r'D:\File\attach'  # 附件所在文件夹
attach_type = ".xlsx"  # 附件后缀名，即类型
addrBook = r'D:\File\测试邮箱组.csv'  # 邮件地址通讯录
year = 2019  # 年份，每次发送前修改
month = 10  # 月份，每次发送前修改

# 根据输入的CSV文件，获取通讯录人名和相应的邮箱地址

def getAddrBook(addrBook):
    with open(addrBook, 'r', encoding='gbk') as addrFile:
        reader = csv.reader(addrFile)
        name = []
        value = []
        addrs = {}
        for row in reader:
            if reader.line_num == 1:  # 跳过表头
                continue;
            else:
                name.append(row[2])
                value.append(row[4])
    addrs = dict(zip(name, value))
    return addrs

# 根据输入的CSV文件，获取通讯录人名和相应的邮箱地址

def getAddrcc(addrBook):
    with open(addrBook, 'r', encoding='gbk') as addrFile:
        reader = csv.reader(addrFile)
        name = []
        value = []
        addrs = {}
        for row in reader:
            if reader.line_num == 1:  # 跳过表头
                continue;
            else:
                name.append(row[2])
                value.append(row[6])
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

mail_content = '''
<p style="font-weight:bold">Dear：</p>
<p>烦请核对附件2019年10月份此处为自变量，会随着月份的变化而修正月提成表。<br />
如下请特别关注：<br />
1、个人和团队进账数据<br />
2、个人当月出单数，上月出单数，退款，计提点及提成<br />
3、下放期内会员进账数据<br />
4、团队当月0单人数<br />
5、团标，团队提成是否发放<br />
</p>
<p>
请您最晚于2019年10月30日此处为自变量，根据每月时间自定义修改下前回复，谢谢支持<br />
</p>
'''

# 加载邮件内容
footer = '''
<p><br /></p>
<p><img src="https://emailalina.oss-cn-beijing.aliyuncs.com/footer.jpg"></p>
<p style="font-size:25px;font-weight:bold">高顿财税学院 | HR Dept</p>
<p style="font-size:15px;font-weight:bold">Alina Zhang</p>
<p>Tel：400-111-0518</p>
<p>Mail：zhangling@goldenfinance.com.cn</p>
<p>Web：www.goldenfinance.com.cn</p>
<p>Add：上海市虹口区花园路171号A5高顿教育</p>
'''


# 添加附件

def addAttch(attach_file):
    att = MIMEBase('application', 'octet-stream')  # 这两个参数不知道啥意思，二进制流文件
    att.set_payload(open(attach_file, 'rb').read())
    # 此时的附件名称为****.xlsx，截取文件名
    att.add_header('Content-Disposition', 'attachment',
                   filename=(attach_file.split("\\")[-1].split(".")[0] + '-' + str(year) + str(month) + '.xlsx'))
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
            subject = att_name + '-' + str(year) + str(month)
            msg['Subject'] = subject  # 设置邮件主题
            person_name = att_name  # subject.split("_")[-1]

            addrs = getAddrBook(addrBook)
            recv_addr = getRecvAddr(addrs, person_name)
            cc_addr = getAddrcc(addrBook)
            cc_recv_addr = getRecvAddr(cc_addr, person_name)
            msg['From'] = formataddr([sender_name, sender_user])  # 设置发件人名称
            # msg['To'] = person_name # 设置收件人名称
            msg['To'] = formataddr([person_name, recv_addr])  # 设置收件人名称
            msg['cc'] = formataddr([person_name,cc_recv_addr]) # 设置抄送人
            msg.attach(MIMEText(mail_content + footer, 'html', 'utf-8'))  # 正文
            attach_file = root + "\\" + attach_file
            att = addAttch(attach_file)
            msg.attach(att)  # 附件
            smtp.sendmail(sender_user, [recv_addr, ], msg.as_string())  # smtp.sendmail(from_addr, to_addrs, msg)
            print("已发送： " + person_name + " <" + recv_addr + ">")
        smtp.quit()


if __name__ == '__main__':
    mailSend()
    anykey = input("请按回车键（Enter）退出程序：")
    if anykey:
        exit(50)
