from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import ssl
import smtplib
import datetime

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

    def set_email_content(self, subject, body, file_paths=None, file_names=None):
        self.msg['From'] = self.email_username
        self.msg['To'] = ', '.join(self.receivers)
        self.msg['Subject'] = subject
        self.msg.attach(MIMEText(body, 'plain'))
        if file_paths and file_names:
            for file_path, file_name in zip(file_paths, file_names):
                att1 = MIMEText(open(file_path + file_name, 'rb').read(), 'base64', 'utf-8')
                att1["Content-Type"] = 'application/octet-stream'
                att1.add_header('Content-Disposition', 'attachment', filename=('gbk', '', file_name))
                self.msg.attach(att1)

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

# Example usage:
if __name__ == '__main__':
    # 初始化 DailyEmailReport 实例
    email_reporter = DailyEmailReport(
        email_host='smtp.xx.com',
        email_port=465,
        email_username='xxx@xx.com',
        email_password='xxx'
    )

    # 添加收件人
    email_reporter.add_receiver('xxx@126.com')
    
    title=""
    text="""
        Hello,

        Here is your daily report.

        [Insert your report content here.]

        Regards,
        Your Name
        """

    # 发送每日报告
    email_reporter.send_daily_report(title,text)
