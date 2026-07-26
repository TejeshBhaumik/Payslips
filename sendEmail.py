
import logging
import os
import base64
import imaplib
import smtplib
import ssl
from email.message import EmailMessage
from email.utils import formatdate
import requests


logger = logging.getLogger(__name__)

SMTP_HOST = os.environ.get("SMTP_HOST", "mail.conacent.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "hr@conacent.com")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "").strip()
SMTP_FROM_EMAIL = os.environ.get("SMTP_FROM_EMAIL", SMTP_USER).strip()
SMTP_DEBUG = os.environ.get("SMTP_DEBUG", "0").strip().lower() in {"1", "true", "yes", "on"}
SMTP_TIMEOUT_SECONDS = int(os.environ.get("SMTP_TIMEOUT_SECONDS", "20"))
SMTP2GO_API_KEY = os.environ.get("SMTP2GO_API_KEY", "").strip()
SMTP2GO_API_URL = os.environ.get("SMTP2GO_API_URL", "https://api.smtp2go.com/v3/email/send")
IMAP_HOST = os.environ.get("IMAP_HOST", "mail.conacent.com")
IMAP_PORT = int(os.environ.get("IMAP_PORT", "993"))
IMAP_SENT_FOLDER = os.environ.get("IMAP_SENT_FOLDER", "Sent")
IMAP_SENT_FOLDER_CANDIDATES = (
    IMAP_SENT_FOLDER,
    "sent",
    "Sent",
    "Sent Items",
    "INBOX.sent",
    "INBOX.Sent",
    "INBOX/sent",
    "INBOX/Sent",
    "[Gmail]/Sent Mail",
)


class EmailRetryNeeded(Exception):
    """Raised when email delivery should pause and retry from the same row."""


class RateLimitExceeded(EmailRetryNeeded):
    """Raised when the SMTP provider rejects sending because of throttling or quota."""


SMTP_RATE_LIMIT_CODES = {421, 451, 452, 454}
SMTP_RATE_LIMIT_TERMS = (
    "rate",
    "limit",
    "quota",
    "throttle",
    "too many",
    "temporarily unavailable",
    "try again later",
    "daily",
    "hourly",
)


def _decode_smtp_error(error):
    message = error.smtp_error
    if isinstance(message, bytes):
        return message.decode("utf-8", errors="replace")
    return str(message)


def _is_smtp_rate_limit(error):
    message = _decode_smtp_error(error).lower()
    return error.smtp_code in SMTP_RATE_LIMIT_CODES or any(term in message for term in SMTP_RATE_LIMIT_TERMS)


def _is_recipient_rate_limit(error):
    for code, message in error.recipients.values():
        text = message.decode("utf-8", errors="replace") if isinstance(message, bytes) else str(message)
        if code in SMTP_RATE_LIMIT_CODES or any(term in text.lower() for term in SMTP_RATE_LIMIT_TERMS):
            return True
    return False


def resolve_smtp_credentials():
    user = SMTP_USER.strip()
    password = SMTP_PASSWORD.strip()
    if not user:
        raise ValueError("SMTP user is required. Set SMTP_USER in the environment.")
    if not password:
        raise ValueError("SMTP password is required. Set SMTP_PASSWORD in the environment.")
    return user, password


def resolve_sender_email(form_sender=None):
    sender = (form_sender or SMTP_FROM_EMAIL).strip()
    if not sender:
        raise ValueError("Sender email is required. Set SMTP_FROM_EMAIL or enter a sender email in the form.")
    return sender


def build_email_message(sender, receiver, subject, text_body, html_body=None, attachments=None):
    message = EmailMessage()
    message["From"] = sender
    message["To"] = receiver
    message["Subject"] = subject
    message["Date"] = formatdate(localtime=True)
    message.set_content(text_body)
    if html_body:
        message.add_alternative(html_body, subtype="html")
    for attachment in attachments or []:
        content_type = attachment.get("content_type", "application/octet-stream")
        maintype, _, subtype = content_type.partition("/")
        if not subtype:
            maintype = "application"
            subtype = "octet-stream"
        message.add_attachment(
            base64.b64decode(attachment["fileblob"]),
            maintype=maintype,
            subtype=subtype,
            filename=attachment["filename"],
        )
    return message


def _decode_imap_value(value):
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return str(value)


def _parse_imap_mailbox_name(raw_mailbox):
    text = _decode_imap_value(raw_mailbox).strip()
    if not text:
        return ""

    if text.endswith('"'):
        start = text.rfind('"', 0, -1)
        if start != -1:
            return text[start + 1 : -1]

    return text.rsplit(" ", 1)[-1].strip('"')


def _format_imap_mailbox_name(folder_name):
    if any(char.isspace() for char in folder_name) or folder_name.startswith("["):
        return '"' + folder_name.replace("\\", "\\\\").replace('"', '\\"') + '"'
    return folder_name


def _find_sent_folder(mailbox):
    unique_candidates = []
    for candidate in IMAP_SENT_FOLDER_CANDIDATES:
        if candidate and candidate not in unique_candidates:
            unique_candidates.append(candidate)

    status, mailboxes = mailbox.list()
    if status != "OK" or not mailboxes:
        logger.warning("IMAP mailbox list failed status=%s data=%s", status, mailboxes)
        return unique_candidates

    discovered = []
    for raw_mailbox in mailboxes:
        text = _decode_imap_value(raw_mailbox)
        folder_name = _parse_imap_mailbox_name(raw_mailbox)
        if not folder_name:
            continue
        if "\\Sent" in text or folder_name.lower() in {"sent", "sent items", "sent mail"}:
            discovered.append(folder_name)
        elif folder_name.lower().endswith((".sent", "/sent")):
            discovered.append(folder_name)

    for folder_name in discovered:
        if folder_name not in unique_candidates:
            unique_candidates.append(folder_name)

    logger.info("IMAP sent-folder candidates=%s", unique_candidates)
    return unique_candidates


def append_to_sent_folder(sender, mailbox_password, message):
    if not mailbox_password:
        logger.info("Skipping IMAP sent-copy append for sender=%s; no mailbox password supplied", sender)
        return False

    try:
        logger.info("IMAP connecting host=%s port=%s sender=%s", IMAP_HOST, IMAP_PORT, sender)
        with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, timeout=SMTP_TIMEOUT_SECONDS) as mailbox:
            mailbox.login(sender, mailbox_password)
            encoded_message = message.as_bytes()
            for folder_name in _find_sent_folder(mailbox):
                status, data = mailbox.append(
                    _format_imap_mailbox_name(folder_name),
                    "\\Seen",
                    imaplib.Time2Internaldate(None),
                    encoded_message,
                )
                if status == "OK":
                    logger.info("IMAP appended sent copy sender=%s folder=%s", sender, folder_name)
                    return True
                logger.warning("IMAP append failed sender=%s folder=%s status=%s data=%s", sender, folder_name, status, data)
    except Exception:
        logger.exception("IMAP sent-copy append failed sender=%s folder=%s", sender, IMAP_SENT_FOLDER)
    return False


def emailDelivery(auth_user, auth_password, message, receiver, bcc=None):
    context = ssl.create_default_context()
    try:
        logger.info("SMTP connecting host=%s port=%s receiver=%s", SMTP_HOST, SMTP_PORT, receiver)
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=SMTP_TIMEOUT_SECONDS) as session:
            session.set_debuglevel(1 if SMTP_DEBUG else 0)
            session.ehlo()
            logger.info("SMTP connected; starting TLS receiver=%s", receiver)
            session.starttls(context=context)
            session.ehlo()

            logger.info("SMTP logging in auth_user=%s receiver=%s", auth_user, receiver)
            session.login(auth_user, auth_password)

            text = message.as_string()
            logger.info("SMTP sending message receiver=%s subject=%s", receiver, message.get("Subject", ""))
            recipients = [receiver]
            if bcc and bcc != receiver:
                recipients.append(bcc)
            session.sendmail(message["From"], recipients, text)
            logger.info("SMTP sent message receiver=%s", receiver)
    except smtplib.SMTPRecipientsRefused as e:
        if _is_recipient_rate_limit(e):
            logger.warning(
                "SMTP provider stopped sending sender=%s receiver=%s recipients=%s",
                auth_user,
                receiver,
                e.recipients,
            )
            raise RateLimitExceeded("Email provider stopped sending. Please retry later.") from e
        logger.exception("SMTP delivery failed auth_user=%s receiver=%s", auth_user, receiver)
        raise
    except smtplib.SMTPResponseException as e:
        if _is_smtp_rate_limit(e):
            smtp_message = _decode_smtp_error(e).strip()
            logger.warning(
                "SMTP provider stopped sending sender=%s receiver=%s code=%s error=%s",
                auth_user,
                receiver,
                e.smtp_code,
                smtp_message,
            )
            raise RateLimitExceeded("Email provider stopped sending. Please retry later.") from e
        logger.exception("SMTP delivery failed auth_user=%s receiver=%s", auth_user, receiver)
        raise
    except TimeoutError as e:
        logger.warning(
            "SMTP connection timed out sender=%s receiver=%s host=%s port=%s",
            auth_user,
            receiver,
            SMTP_HOST,
            SMTP_PORT,
        )
        raise EmailRetryNeeded("Email server connection timed out. Retrying later.") from e
    except Exception:
        logger.exception("SMTP delivery failed auth_user=%s receiver=%s", auth_user, receiver)
        raise


def send_with_smtp2go_api(sender, receiver, subject, text_body, html_body=None, attachments=None, bcc=None):
    payload = {
        "api_key": SMTP2GO_API_KEY,
        "sender": sender,
        "to": [receiver],
        "subject": subject,
        "text_body": text_body,
    }
    if html_body:
        payload["html_body"] = html_body
    if attachments:
        payload["attachments"] = attachments
    if bcc and bcc != receiver:
        payload["bcc"] = [bcc]

    try:
        logger.info("SMTP2GO API sending receiver=%s subject=%s", receiver, subject)
        response = requests.post(SMTP2GO_API_URL, json=payload, timeout=SMTP_TIMEOUT_SECONDS)
    except requests.Timeout as e:
        logger.warning("SMTP2GO API timed out receiver=%s", receiver)
        raise EmailRetryNeeded("Email provider API timed out. Retrying later.") from e
    except requests.RequestException as e:
        logger.exception("SMTP2GO API request failed receiver=%s", receiver)
        raise EmailRetryNeeded("Email provider API request failed. Retrying later.") from e

    if response.status_code == 429:
        logger.warning("SMTP2GO API rate limited receiver=%s response=%s", receiver, response.text[:500])
        raise RateLimitExceeded("Email provider stopped sending. Please retry later.")
    if response.status_code >= 500:
        logger.warning("SMTP2GO API server error status=%s receiver=%s response=%s", response.status_code, receiver, response.text[:500])
        raise EmailRetryNeeded("Email provider API is temporarily unavailable. Retrying later.")
    if not response.ok:
        logger.error("SMTP2GO API rejected message status=%s receiver=%s response=%s", response.status_code, receiver, response.text[:500])
        raise ValueError(f"Email provider rejected the message: HTTP {response.status_code}")

    try:
        result = response.json()
    except ValueError:
        result = {}
    data = result.get("data", {}) if isinstance(result, dict) else {}
    failures = data.get("failures") or []
    if failures:
        logger.error("SMTP2GO API message failures receiver=%s failures=%s", receiver, failures)
        raise ValueError(f"Email provider rejected the message for {receiver}")

    logger.info("SMTP2GO API sent receiver=%s", receiver)


def deliver_email(sender, receiver, subject, text_body, html_body=None, attachments=None, mailbox_password=None):
    bcc = sender
    message = build_email_message(
        sender,
        receiver,
        subject,
        text_body,
        html_body=html_body,
        attachments=attachments,
    )
    if SMTP2GO_API_KEY:
        send_with_smtp2go_api(
            sender,
            receiver,
            subject,
            text_body,
            html_body=html_body,
            attachments=attachments,
            bcc=bcc,
        )
        append_to_sent_folder(sender, mailbox_password, message)
        return

    auth_user, auth_password = resolve_smtp_credentials()
    emailDelivery(auth_user, auth_password, message, receiver, bcc=bcc)
    append_to_sent_folder(sender, mailbox_password, message)


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

    sender = resolve_sender_email(smtp_user)
    receiver = receiver_email

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

    deliver_email(
        sender,
        receiver,
        "Happy holidays from Conacent!",
        body,
        html_body=email_signature,
        mailbox_password=smtp_password,
    )

    







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
    sender = resolve_sender_email(smtp_user)
    receiver = email
    deliver_email(
        sender,
        receiver,
        f" {month}",
        body,
        attachments=[
            {
                "filename": pdf_name,
                "fileblob": base64.b64encode(pdf_bytes).decode("ascii"),
                "content_type": "application/pdf",
            }
        ],
        mailbox_password=smtp_password,
    )
