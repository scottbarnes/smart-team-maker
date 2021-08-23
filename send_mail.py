# Larhgely from https://realpython.com/python-send-email/#option-2-setting-up-a-local-smtp-server

import smtplib, ssl, sys
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
# import click
from jinja2 import Environment, FileSystemLoader, select_autoescape
from jinja2.exceptions import TemplateNotFound

def send_mail(
    mail_server,
    port,
    to,
    sender_email,
    password,
    subject,
    template,
    body,
    cc=None,
    bcc=None,
    **kwargs,
):
    """ Send email. """

    # Set up the Jinja2 template.
    template_loader = FileSystemLoader(searchpath='./email_templates')
    template_env = Environment(loader=template_loader)
    # template = 'initial_team_email.jinja2'
    try:
        template = template_env.get_template(template)
    except TemplateNotFound as e:
        print(f'\nError: template name "{e}" could not be found.\nEnsure you\'ve entered a template FILENAME From the template directory and not a template PATH.')
        sys.exit(1)
    body = template.render(**kwargs)

    print(f'to: {to}')
    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = sender_email
    message["To"] = ', '.join(to)
    print(f'to: {message["To"]}')

    # # Create the plain-text and HTML version of your message
    # text = """\
    # This section is not used."""
    # html = """\
    # <html>
    #   <body>
    #     <p>Hi,<br>
    #        This section is not used<br>
    #     </p>
    #   </body>
    # </html>
    # """

    # Turn these into plain/html MIMEText objects
    # part1 = MIMEText(text, "plain")
    # part2 = MIMEText(html, "html")

    email_content = MIMEText(body, "html")

    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(email_content)
    # message.attach(part2)

    # Create secure connection with server and send email
    context = ssl.create_default_context()
    print(f'mail_server: {mail_server}\nport: {port}')
    with smtplib.SMTP_SSL(mail_server, port, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(
            sender_email, to, message.as_string()
        )
