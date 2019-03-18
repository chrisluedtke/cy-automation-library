from configparser import ConfigParser
from pathlib import Path

import smtplib

from .config import OKTA_USER, OKTA_PASS

def send_email(to_addrs, subject, body):
    """ Sends an email to one or multple `to_addrs` email addresses.
    """
    username = OKTA_USER
    password = OKTA_PASS
    if '@cityyear.org' not in username:
        username += '@cityyear.org'

    mailserver = smtplib.SMTP('smtp.office365.com', 587)
    mailserver.ehlo()
    mailserver.starttls()
    mailserver.login(username, password)
    msg = f'''From: {username}\nSubject: {subject}\n\n{body}'''
    mailserver.sendmail(from_addr=username, to_addrs=to_addrs, msg=msg)
    mailserver.quit()
