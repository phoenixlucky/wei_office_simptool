from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import ssl
import smtplib
import datetime
class eSend(object):
    """
    新邮件系统,可群发,群带附件
    """
    def __init__(self,sender=None,receiver=None,username=None,password=None,smtpserver='smtp.126.com'):
        self.sender = sender
        self.receiver = receiver
        self.username = username
        self.password = password
        self.smtpserver = smtpserver

    def send_email(self, subject,e_content, file_paths, file_names):
        try:
            message = MIMEMultipart()
            message['From'] = self.sender  # 发送
            message['To'] = ",".join(self.receiver)  # 收件
            message['Subject'] = Header(subject, 'utf-8')
            message.attach(MIMEText(e_content, 'plain', 'utf-8'))  # 邮件正文

            # 构造附件群
            for file_path,file_name in zip(file_paths,file_names):
                print(file_name,file_path)
                att1 = MIMEText(open(file_path + file_name, 'rb').read(), 'base64', 'utf-8')
                att1["Content-Type"] = 'application/octet-stream'
                att1.add_header('Content-Disposition', 'attachment', filename=('gbk', '', file_name))
                message.attach(att1)

            # 执行
            smtpSsl=smtplib.SMTP_SSL(self.smtpserver)
            smtpSsl.connect(self.smtpserver,465)  # 连接服务器
            smtpSsl.login(self.username, self.password)  # 登录
            smtpSsl.sendmail(self.sender, self.receiver, message.as_string())  # 发送
            smtpSsl.quit()
            print("The email with file_names has been send!")
        except Exception as e:
            print(e)
            pass

class DailyEmailReport:
    def __init__(self, email_host, email_port, email_username, email_password):
        self.email_host = email_host
        self.email_port = email_port
        self.email_username = email_username
        self.email_password = email_password
        self.receivers = []
        self.msg = MIMEMultipart()

    def add_receiver(self, receiver_email):
        self.receivers.append(receiver_email)

    def set_email_content(self, subject, body):
        self.msg['From'] = self.email_username
        self.msg['To'] = ', '.join(self.receivers)
        self.msg['Subject'] = subject
        self.msg.attach(MIMEText(body, 'plain'))

    def send_email(self):
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(self.email_host, self.email_port, context=context) as server:
            server.login(self.email_username, self.email_password)
            server.sendmail(self.email_username, self.receivers, self.msg.as_string())
            print('邮件发送成功！')

    def send_daily_report(self,title,text):
        subject = f'{title} - {datetime.date.today()}'
        body = text
        self.set_email_content(subject, body)
        self.send_email()
