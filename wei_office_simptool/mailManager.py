from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
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
