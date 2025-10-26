from dotenv import load_dotenv
load_dotenv()
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

sender = os.getenv("MAIL_USERNAME")
password = os.getenv("MAIL_PASSWORD")
receiver = os.getenv("TEACHER_EMAIL", sender)

print("Gửi từ:", sender)
print("Mật khẩu (ẩn):", "*" * len(password) if password else "MISSING")

if not password:
    print("❌ Thiếu MAIL_PASSWORD!")
    exit(1)

msg = MIMEMultipart()
msg['From'] = sender
msg['To'] = receiver
msg['Subject'] = "✅ Test hệ thống gửi email"

body = "Nếu bạn thấy email này, hệ thống đã cấu hình thành công!"
msg.attach(MIMEText(body, 'plain', 'utf-8'))

try:
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(sender, password)
    server.sendmail(sender, receiver, msg.as_string())
    server.quit()
    print("📧 Gửi thành công! Vui lòng kiểm tra hộp thư (kể cả Spam).")
except Exception as e:
    print(f"❌ Lỗi gửi email: {e}")