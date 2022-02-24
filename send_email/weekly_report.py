"""
@author:maohui
@time:2022/2/24 11:45
"""

import smtplib  # 发送
from smtplib import SMTP_SSL  # 加密邮件内容
from email.mime.text import MIMEText  # 构造邮件的正文
from email.mime.application import MIMEApplication  # 添加附件
from email.header import Header
from email.mime.multipart import MIMEMultipart  # 邮件的主体

# 输入email地址和口令
from_addr = "409788696@qq.com"
authcodes = "fzzrntnmwuhgbgbb"
# 输入收件人地址：
to_addr = "409788696@qq.com"
# 输入抄送人地址：
cc_addr = "409788696@qq.com"
# 输入SMTP服务器地址
smtp_server = "smtp.qq.com"

# 输入主题
mail_title = "周报"
# 输入邮件内容
mail_content = "尊敬的领导：<p>您们好！</p><p>以下是我的周报，请查收，如有不足，请提出" \
               "您宝贵的意见，我将在以后的工作中即时改进，谢谢！</p>"

# 初始化对象
msg = MIMEMultipart()  # 初始化邮件主体
msg['Subject'] = mail_title  # 放入标题
msg['From'] = from_addr  # 放入发送人
msg['To'] = to_addr  # 放入收件人
msg['Cc'] = cc_addr  # 放入抄送人

# 发送正文
msg.attach(MIMEText(mail_content, 'html', 'utf-8'))

# 打开附件
xlsx = MIMEApplication(open('C:\周报\周报_毛辉_2022-02-25.xlsx', 'rb').read())
# 添加一个头部
xlsx.add_header('Content-Dispositon', 'attachment', filename='周报-毛辉-2022-02-25')
# 将附件一起发送
msg.attach(xlsx)

smtp = SMTP_SSL(smtp_server, 465)  # 连接发送的邮箱服务器
smtp.login(from_addr, authcodes)  # 登录发送的邮箱
smtp.sendmail(from_addr, to_addr, msg.as_string())
smtp.quit()

# try:
#     smtp = SMTP_SSL(smtp_server)
#     smtp.set_debuglevel(1)
#     smtp.ehlo(smtp_server)
#     print("邮件发送成功")
# except smtplib.SMTPException:
#     print("无法发送邮件")
