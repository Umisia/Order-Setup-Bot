import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import config


def send_internal_email(subject, email_to, content):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = config.from_email_username
    msg['To'] = email_to
    if content == "send logs":
        msg.attach(MIMEText("Program stopped working. Check attached logs.", 'html'))
        part = MIMEBase('application', "octet-stream")
        with open(config.logs_path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="order_setup.log"')
        msg.attach(part)
    else:
        msg.attach(MIMEText(content, 'html'))
    mail = smtplib.SMTP('smtp.gmail.com', 587)
    mail.ehlo()
    mail.starttls()
    mail.login(config.from_email_username, config.from_email_pswd)
    mail.sendmail(msg['From'], msg['To'], msg.as_string())
    mail.quit()