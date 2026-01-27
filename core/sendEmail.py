import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

from dotenv import load_dotenv
import os
load_dotenv()


def send_email(subject, body, sender, recipients, password, pdf_path):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ', '.join(recipients)

    # Cuerpo del correo
    msg.attach(MIMEText(body, 'plain'))

    # Adjuntar PDF
    pdf_path = Path(pdf_path)

    with open(pdf_path, "rb") as f:
        part = MIMEBase("application", "pdf")
        part.set_payload(f.read())

    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f'attachment; filename="{pdf_path.name}"'
    )

    msg.attach(part)

    # Envío
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp_server:
        smtp_server.login(sender, password)
        smtp_server.sendmail(sender, recipients, msg.as_string())

    print("Message sent with attachment!")

if __name__=="__main__":
    subject = "Email Subject"
    body = "This is the body of thetexmessage"
    sender = "99lotermin@gmail.com"
    recipients = ["99lotermin1@gmail.com"]
    password = os.getenv("PASSWORD")
    pdf_path = r"C:\Users\Loren Otermin\Downloads\erreserba.pdf"
    send_email(subject, body, sender, recipients, password, pdf_path)
