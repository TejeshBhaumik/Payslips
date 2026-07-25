
import logging
import os
import smtplib
import threading
import time
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


logger = logging.getLogger(__name__)

SMTP_HOST = os.environ.get("SMTP_HOST", "mail.conacent.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "hr@conacent.com")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "").strip()
SMTP_DEBUG = os.environ.get("SMTP_DEBUG", "0").strip().lower() in {"1", "true", "yes", "on"}
MAX_EMAILS_PER_BATCH = int(os.environ.get("MAX_EMAILS_PER_BATCH", "30"))
RATE_LIMIT_COOLDOWN_SECONDS = int(os.environ.get("RATE_LIMIT_COOLDOWN_SECONDS", "300"))

_send_lock = threading.Lock()
_sent_in_batch = 0


def reserve_send_slot():
    global _sent_in_batch
    with _send_lock:
        if _sent_in_batch > 0 and _sent_in_batch % MAX_EMAILS_PER_BATCH == 0:
            logger.info(
                "Sent %d emails; sleeping %d seconds before continuing",
                _sent_in_batch,
                RATE_LIMIT_COOLDOWN_SECONDS,
            )
            time.sleep(RATE_LIMIT_COOLDOWN_SECONDS)

        _sent_in_batch += 1


def resolve_smtp_credentials(smtp_user=None, smtp_password=None):
    user = (smtp_user or SMTP_USER).strip()
    password = (smtp_password or SMTP_PASSWORD).strip()
    if not user:
        raise ValueError("SMTP user is required. Set SMTP_USER or provide smtpEmail in the form.")
    if not password:
        raise ValueError("SMTP password is required. Set SMTP_PASSWORD or provide smtpPassword in the form.")
    return user, password


def emailDelivery(smtp_user, smtp_password, message, receiver):
    reserve_send_slot()
    logger.info(
        "SMTP send starting host=%s port=%s sender=%s receiver=%s subject=%s",
        SMTP_HOST,
        SMTP_PORT,
        smtp_user,
        receiver,
        message.get("Subject", ""),
    )

    context = ssl.create_default_context()
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as session:
            session.set_debuglevel(1 if SMTP_DEBUG else 0)
            logger.info("SMTP connection established")
            session.ehlo()
            session.starttls(context=context)
            logger.info("SMTP STARTTLS complete")
            session.ehlo()

            logger.info("SMTP login attempt sender=%s", smtp_user)
            session.login(smtp_user, smtp_password)
            logger.info("SMTP login success sender=%s", smtp_user)

            text = message.as_string()
            logger.info("SMTP sendmail starting receiver=%s", receiver)
            session.sendmail(smtp_user, receiver, text)
            logger.info("SMTP sendmail success receiver=%s", receiver)
    except Exception:
        logger.exception("SMTP delivery failed sender=%s receiver=%s", smtp_user, receiver)
        raise


def sendEmailWithImage(
    person_name,
    img,
    sender,
    sender_title,
    company,
    com_address,
    ph_number,
    com_email,
    receiver_email,
    smtp_user=None,
    smtp_password=None,
):
    
    body = f''' Dear {person_name},
                Happy holidays and a fantastic New Year to you!
                Wishing you joy, peace, and prosperity in the coming year
                '''

    # put your email here
    sender, password = resolve_smtp_credentials(smtp_user, smtp_password)
    # put the email of the receiver here
    receiver = receiver_email

    #Setup the MIME
    message = MIMEMultipart()
    message['From'] = sender
    message['To'] = receiver
    message["Bcc"] = sender
    message['Subject'] = 'Happy holidays from Conacent!'

    #read text from file if shipped 
    email_signature = '''
            <!DOCTYPE html>
            <html>
            <head>
            <title>Email Signature</title>
            <style>
                body {
                font-family: Arial, sans-serif;
                line-height: 1.6;
                margin: 0;
                padding: 0;
                }
                .email-signature {
                max-width: 500px;
                margin: 0 auto;
                font-size: 16px;
                color: #333;
                border-bottom: 2px solid #ccc;
                padding: 20px;
                }
                .name {
                font-size: 18px;
                font-weight: bold;
                margin-bottom: 5px;
                }
                .title {
                font-size: 16px;
                margin-bottom: 5px;
                }
                .company {
                font-size: 16px;
                margin-bottom: 5px;
                }
                .address {
                font-size: 16px;
                margin-bottom: 5px;
                }
                .phone {
                font-size: 16px;
                margin-bottom: 5px;
                }
                .website {
                font-size: 16px;
                margin-bottom: 5px;
                text-decoration: none;
                color: #007bff;
                }
            </style>
            </head>
            <body>

            <div class="email-signature">
            <p class="name">Pranabesh Bhaumik</p>
            <p class="title">Director</p>
            <p class="company">Conacent Consulting Pvt Ltd</p>
            <p class="address">CF-90 Salt Lake, Sector 1<br>Kolkata - 700064, INDIA</p>
            <p class="phone">Ph: +91 98300 79710</p>
            <p class="website"><a href="http://www.conacent.com" target="_blank">www.conacent.com</a></p>
            </div>

            </body>
            </html>
    '''


    message.attach(MIMEText(body, 'plain'))
    message.attach(MIMEText(email_signature, 'html'))
    emailDelivery(sender, password, message, receiver)

    







def sendEmailWithPDF(
    pdf_bytes,
    pdf_name,
    email,
    month,
    person_name,
    email_file_body,
    is_pan,
    smtp_user=None,
    smtp_password=None,
):

    body = f''' Dear {person_name}, 
                Please find attached the {email_file_body} {month}.
                {is_pan}
                HR'''
    # put your email here
    sender, password = resolve_smtp_credentials(smtp_user, smtp_password)
    # put the email of the receiver here
    receiver = email
      
    #Setup the MIME
    message = MIMEMultipart()
    message['From'] = sender
    message['To'] = receiver
    message["Bcc"] = sender
    message['Subject'] = f' {month}'

    message.attach(MIMEText(body, 'plain'))

    pdfname = pdf_name

    payload = MIMEBase('application', 'octate-stream', Name=pdfname)
    payload.set_payload(pdf_bytes)
    logger.info("Prepared PDF attachment name=%s size_bytes=%d", pdfname, len(pdf_bytes))

    # enconding the binary into base64
    encoders.encode_base64(payload)

    # add header with pdf name
    payload.add_header('Content-Disposition', 'attachment', filename=pdfname)
    message.attach(payload)
    logger.info("Dispatching PDF email recipient=%s subject=%s", receiver, message.get("Subject", ""))
    emailDelivery(sender, password, message, receiver)
