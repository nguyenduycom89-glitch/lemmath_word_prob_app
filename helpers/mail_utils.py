import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def send_teacher_notification(student_name, submission_id, base_url, config, teacher_email):
    """Gửi email thông báo khi học sinh nộp bài"""
    subject = f"🧮 Học sinh {student_name} vừa nộp bài #{submission_id}"
    body = f"""
    Xin chào thầy/cô,

    Học sinh **{student_name}** vừa nộp bài mới (ID: {submission_id}).
    Xem chi tiết tại: {base_url}teacher

    — Hệ thống Toán học thông minh 🧠
    """
    msg = MIMEMultipart()
    msg["From"] = config.get("MAIL_USERNAME")
    msg["To"] = teacher_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))

    try:
        with smtplib.SMTP(config["MAIL_SERVER"], config["MAIL_PORT"]) as server:
            if config["MAIL_USE_TLS"]:
                server.starttls()
            server.login(config["MAIL_USERNAME"], config["MAIL_PASSWORD"])
            server.send_message(msg)
    except Exception as e:
        print("❌ Không gửi được email:", e)
