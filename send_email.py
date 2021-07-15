import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from smtpd import COMMASPACE

class SendEmailByGoogleMail:
    def __init__(self, subject, username, password, receivers:list):
        # 初始化帳號
        self.user_account = {'username': username, 'password': password}
        # 初始化主旨
        self.subject = subject
        # 設置gmail服務器
        self.smtp_server = 'smtp.gmail.com'
        # SSL端口
        self.port = 465
        # 初始化寄件人姓名
        self.sender = ''
        # 初始化收件人郵箱
        self.receivers = receivers

    def send_mail(self, way, content, files):
        msg_root = MIMEMultipart()
        # 建立附件列表
        if files is not None:
            for file in files:
                file_name = file.split("/")[-1]
                att = MIMEText(open(file, 'rb').read(), 'base64', 'utf-8')
                att["Content-Type"] = 'application/octet-stream'
                att["Content-Disposition"] = 'attachment; filename=%s' % file_name
                msg_root.attach(att)

        # 建立郵件主旨
        msg_root['Subject'] = self.subject
        # 接收者暱稱
        msg_root['To'] = ';'.join(self.receivers)
        # 郵件內文
        if way == 'common':
            msg_root.attach(MIMEText(content, 'plain', 'utf-8'))
        elif way == 'html':
            msg_root.attach(MIMEText(content, 'html', 'utf-8'))
        smtp = smtplib.SMTP_SSL(self.smtp_server, self.port)
        # 開啟Debug
        smtp.set_debuglevel(True)
        smtp.ehlo()
        smtp.login(self.user_account['username'], self.user_account['password'], initial_response_ok=False)
        smtp.auth_plain()
        smtp.sendmail(self.sender, self.receivers, msg_root.as_string())
        print("郵件發送成功")

 
path = input("請輸入檔案路徑(csv): ")
df = pd.read_csv(fr"{path}")
print(df)

column = input("請輸入email資料所在欄位: ")
email_df = df[(df[column].notna() == True)]
print(email_df)

receivers = list(email_df[column])
print(receivers)

sender = input("請輸入你的完整gmail帳號(xxx@gmail.com): ")
password = input("請輸入你的應用程式密碼(16位): ")
test_receiver = input("請輸入測試用接收帳號(如有多個帳號請以空白鍵隔開): ").split()
mail_subject = input("請輸入主旨: ")
mail_content = input("請輸入內文: ")
mail_test_sending = input("是否需要測試寄送(y: 只寄送到測試用接收帳號，n: 寄送到正式收件人帳號): ")

if mail_test_sending == "Y".casefold():
    final_receivers = test_receiver
else:
    final_receivers = receivers
    
sebg = SendEmailByGoogleMail(subject=mail_subject, 
                             username=sender, 
                             password=password,
                             receivers=final_receivers)
sebg.send_mail('common', content=mail_content, files=None)
