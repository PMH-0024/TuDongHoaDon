import smtplib
import ssl
from email.message import EmailMessage

EMAIL_GUI = "hp859869@gmail.com"
MAT_KHAU_UNG_DUNG = "wvqo nsxe oels rqmz"  # ghi đúng chuỗi mật khẩu ứng dụng
EMAIL_NHAN = "hp859869@gmail.com"

msg = EmailMessage()
msg["From"] = EMAIL_GUI
msg["To"] = EMAIL_NHAN
msg["Subject"] = "Test Email từ Python"
msg.set_content("✅ Gửi thành công nếu bạn thấy email này!")

context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(EMAIL_GUI, MAT_KHAU_UNG_DUNG)
    server.send_message(msg)

print("✅ Gửi thành công!")
