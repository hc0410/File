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
sender_user = '1206064486@qq.com' # 发件人邮箱
sender_pwd = 'ntidvudghbqgfjhg' # 邮箱密码（有些为授权码）
sender_name = u'Alina Zhang'  # 发件人

attach_path = r'D:\File\attach'  # 附件所在文件夹
attach_type = ".xlsx"  # 附件后缀名，即类型
addrBook = r'D:\File\receiver.csv'  # 邮件地址通讯录

year = str(2019)  # 年份，每次发送前修改
attach_date = '2019年10月' # 邮件主题后面的日期和附件后面的日期，每次发送前修改
check_date = '2019年10月' # 邮件正文核对提成表的日期，每次发送前修改
reply_date = '2019年11月10日' # 邮件正文最晚回复日期，每次发送前修改

# 获取对应邮箱的参数
a = 2   # 获取分公司的名称
b = 4   # 获取分公司对应政委（收件人）的邮箱
c = 6   # 获取分公司对应分总（收件人）的邮箱
d = 7   # 获取分公司对应抄送人A的邮箱
e = 8   # 获取分公司对应抄送人B的邮箱
f = 9   # 获取分公司对应抄送人C的邮箱
g = 10  # 获取分公司对应抄送人D的邮箱

# 加载邮件内容

content = '''
<p style="font-weight:bold">Dear：</p>
<p>烦请核对附件check_date提成表。<br />
如下请特别关注：<br />
1、个人和团队进账数据<br />
2、个人当月出单数，上月出单数，退款，计提点及提成<br />
3、下放期内会员进账数据<br />
4、团队当月0单人数<br />
5、团标，团队提成是否发放<br />
</p>
<p>
请您最晚于reply_date前回复，谢谢支持<br />
</p>
'''
mail_content = content.replace('check_date',check_date).replace('reply_date',reply_date)

# 加载邮件内容
footer = '''
<p><br /></p>
<p><img src="https://emailalina.oss-cn-beijing.aliyuncs.com/footer.jpg"></p>
<p style="font-size:25px;font-weight:bold">高顿财税学院 | HR Dept</p>
<p style="font-size:15px;font-weight:bold">Alina Zhang</p>
<p>
 <tr>
    <strong>Tel：</strong>
    <th>400-111-0518</th>
 </tr>
</p>
<p>
 <tr>
    <strong>Mail：</strong>
    <th>zhangling@goldenfinance.com.cn</th>
 </tr>
</p>
<p>
 <tr>
    <strong>Web：</strong>
    <a href="www.goldenfinance.com.cn">www.goldenfinance.com.cn</a>
 </tr>
</p>
<p>
 <tr>
    <strong>Add：</strong>
    <th>上海市虹口区花园路171号A5高顿教育</th>
 </tr>
</p>
'''

# 获取邮箱
def getAddrBook(addrBook,a,b):
    with open(addrBook, 'r', encoding='gbk') as addrFile:
        reader = csv.reader(addrFile)
        name = []
        value = []
        addrs = {}
        for row in reader:
            if reader.line_num == 1:  # 跳过表头
                continue;
            else:
                name.append(row[a])
                value.append(row[b])
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

# 添加附件

def addAttch(attach_file):
    att = MIMEBase('application', 'octet-stream')  # 这两个参数不知道啥意思，二进制流文件
    att.set_payload(open(attach_file, 'rb').read())
    # 此时的附件名称为****.xlsx，截取文件名
    att.add_header('Content-Disposition', 'attachment',
                   filename=('gbk','',attach_file.split("\\")[-1].split(".")[0] + '-' + attach_date + attach_type))
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
            subject = att_name + '-' + attach_date
            msg['Subject'] = subject  # 设置邮件主题
            person_name = att_name
            addrs_zhengwei = getAddrBook(addrBook,2,4)
            addrs_fenzong = getAddrBook(addrBook,2,6)
            recv_zhengwei = getRecvAddr(addrs_zhengwei, person_name)
            recv_fenzong = getRecvAddr(addrs_fenzong, person_name)
            addrs_cca = getAddrBook(addrBook,2,7)
            addrs_ccb = getAddrBook(addrBook,2,8)
            addrs_ccc = getAddrBook(addrBook,2,9)
            addrs_ccd = getAddrBook(addrBook,2,10)
            recv_cca = getRecvAddr(addrs_cca, person_name)
            recv_ccb = getRecvAddr(addrs_ccb, person_name)
            recv_ccc = getRecvAddr(addrs_ccc, person_name)
            recv_ccd = getRecvAddr(addrs_ccd, person_name)
            
            msg['From'] = formataddr([sender_name, sender_user])  # 设置发件人名称
            msg['To'] = recv_zhengwei + ';' + recv_fenzong # 设置收件人名称
            msg['cc'] = recv_cca + ';' + recv_ccb + ';' + recv_ccc + ';' + recv_ccd # 设置抄送人
            msg.attach(MIMEText(mail_content + footer, 'html', 'utf-8'))  # 正文
            toaddrs = [recv_zhengwei] + [recv_fenzong] + [recv_cca] + [recv_ccb] + [recv_ccc] + [recv_ccd]
            attach_file = root + "\\" + attach_file
            att = addAttch(attach_file)
            msg.attach(att)  # 附件
            smtp.sendmail(sender_user,toaddrs,msg.as_string())  # smtp.sendmail(from_addr, to_addrs, msg)
            print(person_name + "已发送完成；" + "   收件人： " + person_name + " <" + recv_zhengwei + ';' + recv_fenzong + ">" + "   抄送人：" + "<" + recv_cca + ';' + recv_ccb + ';' + recv_ccc + ';' + recv_ccd +">" + "   附件：" + person_name +".xlsx")
        smtp.quit()

if __name__ == '__main__':
    mailSend()
    anykey = input("请按回车键（Enter）退出程序：")
    if anykey:
        exit(50)
