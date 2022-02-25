"""
@author:maohui
@time:2022/2/24 11:45
"""

import smtplib  # 发送
import time
from smtplib import SMTP_SSL  # 加密邮件内容
from email.mime.text import MIMEText  # 构造邮件的正文
from email.mime.application import MIMEApplication  # 添加附件
from email.mime.multipart import MIMEMultipart  # 邮件的主体
import datetime

# 输入email地址和口令
from_addr = "maohui@well-healthcare.com"
authcodes = "HY8KcFe3wb8GZ6Wu"
# 输入收件人地址：
to_addr = "409788696@qq.com"
# 输入抄送人地址：
cc_addr = "maohui@well-healthcare.com"
# 输入SMTP服务器地址
smtp_server = "smtp.exmail.qq.com"

# 输入主题
mail_title = '周报_毛辉_'+str(datetime.date.today())
print(mail_title)
# 输入邮件内容
mail_content = "尊敬的领导们：" \
               "<p>&nbsp&nbsp&nbsp&nbsp您们好！</p>" \
               "<p>&nbsp&nbsp&nbsp&nbsp附件为我2022-02-21至"+str(datetime.date.today())+"的工作周报，请您查阅，如有不足需改进的地方，请您提出宝贵意见，我将在日后的工作中及时改进。谢谢！</p>"
file_path='c:\周报\\'+mail_title+'.xlsx'
print(file_path)
# 打开附件
xlsx = MIMEApplication(open(file_path, 'rb').read())
# 添加一个头部
xlsx.add_header('Content-Disposition', 'attachment', filename=mail_title+'.xlsx')


# 初始化对象
msg = MIMEMultipart()  # 初始化邮件主体
msg['Subject'] = mail_title  # 放入标题
msg['From'] = from_addr  # 放入发送人
msg['To'] = to_addr  # 放入收件人
msg['Cc'] = cc_addr  # 放入抄送人
# 发送正文
msg.attach(MIMEText(mail_content, 'html', 'utf-8'))
# 将附件一起发送
msg.attach(xlsx)

try:
    smtp = SMTP_SSL(smtp_server, 465)  # 连接发送的邮箱服务器
    smtp.login(from_addr, authcodes)  # 登录发送的邮箱
    smtp.sendmail(from_addr, to_addr.split(",") + cc_addr.split(","), msg.as_string())
    print("邮件发送成功")
    smtp.quit()
except smtplib.SMTPException:
    print('无法发送邮件')

