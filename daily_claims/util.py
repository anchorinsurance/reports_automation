# -*- coding: utf-8 -*-
"""
Created on Tue Jun 24 08:02:23 2017

@author: Surendra Pepakayala
"""

from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
import smtplib
import datetime
import textwrap

from configparser import ConfigParser

config = ConfigParser()
config.read('./config.ini')

def has_attribute(data, attribute):
    return attribute in data and data[attribute] is not None

def send_mail(send_from, send_to, send_cc, subject, text, files=None):

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject
    msg['CC'] = COMMASPACE.join(send_cc)
    msg.attach(MIMEText(text))

    for file in files or []:
        with open(file, "rb") as fil:
            part = MIMEApplication(fil.read(), Name=basename(file))
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(file)
        msg.attach(part)

    smtp = smtplib.SMTP(config.get('email', 'smtp_host'))
    #smtp.login(config.get('email', 'smtp_username'), config.get('email', 'smtp_password'))
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.close()

def log(log_file, msg, indent=True):

    f = open(log_file, "a")
    msg = msg + "\n"
    if (indent == True):
        msg = textwrap.indent(msg, "\t")
    msg = str(datetime.datetime.now()) + "\n" + msg

    try:
        f.writelines(msg)
    finally:
        f.close()
