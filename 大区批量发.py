# -*- coding: utf-8 -*-
"""
Created on Tue Jun 18 13:59:29 2019

@author: Administrator
"""

import smtplib
from email.mime.text import MIMEText
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email import encoders
from email.mime.base import MIMEBase
 
# 第三方 SMTP 服务
mail_host="smtp.exmail.qq.com"  #设置服务器
mail_user="shujufenxi@goldenfinance.com.cn"    #用户名
mail_pass=""   #密码
IO = r'C://Users/Administrator/Desktop/其他信息/各大区邮件.xlsx'# 收件人。抄送人邮箱
xl = pd.ExcelFile('C://Users/Administrator/Desktop/其他信息/各大区邮件.xlsx')
bbz= xl.sheet_names
sheetx=pd.read_excel(IO,sheetname="日常",header=0)
a="-8月"     #每次发送修改！
for i in range(0,41):
    


#1.大区

    attach_path = r'C://Users/Administrator/Desktop/大区日报数据'   # 附件所在文件夹路径
    attach_type = ".xlsm"      # 附件后缀名，即类型
#addrBook = r'C:\Users\user\Desktop\MailMaster V1.1\邮箱联系人表单.csv'

    sender = 'shujufenxi@goldenfinance.com.cn'
    sheet1=pd.read_excel(IO,sheetname=i,header=0)
    receivers = sheet1['邮箱'].values.tolist()  # 收邮件人邮箱！！！！！
    cc=sheet1[sheet1['邮箱.1'].isna()==False]['邮箱.1'].values.tolist()# 抄送人邮箱！！！！！
    cc1=sheetx['邮箱'].values.tolist()

    message= MIMEMultipart()
# 构造附件1，传送当前目录下的文件
    att1 = MIMEText(open('C://Users/Administrator/Desktop/大区日报数据/'+bbz[i]+a+'.xlsx', 'rb').read(), 'base64', 'utf-8')
# 这里的filename可以任意写，写什么名字，邮件中显示什么名字
    att1.add_header('content-disposition', 'attachment', filename=bbz[i]+a+".xlsx")
    message.attach(att1)

# html网页内容
    mail_msg = """
    <p></p>
    <p>&emsp;<img src="cid:image2"></p>
    <p>销售运营中心-数据分析小组</p>
    <p>Mail:  shujufenxi@goldenfinance.com.cn</p>
    <p>Web: www.goldenfinance.com.cn</p>
    <p>Add: 上海市虹口区花园路171号A5高顿教育</p>
    <p>Please consider the environment before printing this email.</p>
    """ 

# 图片2
    fp2 = open('C://Users/Administrator/Desktop/大区日报数据/高顿.jpg', 'rb')
    msgImage1 = MIMEImage(fp2.read())
    fp2.close()
# 定义图片 ID，在 HTML 文本中引用
 
    msgImage1.add_header('Content-ID', '<image2>')
    message.attach(msgImage1)






    message.attach(MIMEText(mail_msg,'html','utf-8'))

    message['From'] = Header("数据分析小组", 'utf-8')
#收件人
    message['To'] =  Header("；".join(sheet1['收件人'].values.tolist()), 'utf-8')      #收件人姓名！！！！！！
#抄送
    message['Cc'] = Header("；".join(sheet1[sheet1['邮箱.1'].isna()==False]['抄送人'].values.tolist()+sheetx['收件人'].values.tolist()),'utf-8')
#邮件名称
    subject = bbz[i]+a      # 名称！！！！！
    message['Subject'] = Header(subject, 'utf-8')
    
    toaddrs = receivers+cc+cc1


    try:
        smtpObj = smtplib.SMTP_SSL("smtp.exmail.qq.com", port=465) 
        smtpObj.login(mail_user,mail_pass)  
        smtpObj.sendmail(sender, toaddrs, message.as_string())
        print ("邮件发送成功")
    except smtplib.SMTPException:
        print (bbz[i]+"Error: 无法发送邮件")
    
    
    
