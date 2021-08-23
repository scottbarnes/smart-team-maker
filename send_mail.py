# Largely from https://realpython.com/python-send-email/#option-2-setting-up-a-local-smtp-server

import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import click

def send_mail(
    mail_server,
    port,
    to,
    sender_email,
    password,
    subject,
    body,
    cc=None,
    bcc=None
):
    """ Send email. """
    print(f'to: {to}')
    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = sender_email
    message["To"] = ', '.join(to)
    print(f'to: {message["To"]}')

    # Create the plain-text and HTML version of your message
    text = """\
    Hi,
    How are you?
    Real Python has many great tutorials:
    www.realpython.com"""
    html = """\
    <html>
      <body>
        <p>Hi,<br>
           How are you?<br>
           <a href="http://www.realpython.com">Real Python</a> 
           has many great tutorials.
        </p>
      </body>
    </html>
    """

    # Turn these into plain/html MIMEText objects
    part1 = MIMEText(text, "plain")
    part2 = MIMEText(html, "html")

    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(part1)
    message.attach(part2)

    # Create secure connection with server and send email
    context = ssl.create_default_context()
    print(f'mail_server: {mail_server}\nport: {port}')
    with smtplib.SMTP_SSL(mail_server, port, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(
            sender_email, to, message.as_string()
        )
