# -*- coding: utf-8 -*-
"""
Created on Tue Nov 4 08:02:23 2017

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

def has_attribute(data, attribute):
    return attribute in data and data[attribute] is not None

def send_mail(host, user, pwd,
              send_from, send_to, send_cc, subject, text,
              files=None, type='text'):

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject
    msg['CC'] = COMMASPACE.join(send_cc)
    msg.attach(MIMEText(text, type, 'utf-8'))

    for file in files or []:
        with open(file, "rb") as fil:
            part = MIMEApplication(fil.read(), Name=basename(file))
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(file)
        msg.attach(part)

    smtp = smtplib.SMTP(host)
    if user.strip() != "":
        smtp.login(user, pwd)
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

def get_db_connection_string(db_host, db_name, db_trusted, db_user, db_pwd):

    connect_string = "Driver={SQL Server Native Client 11.0};"
    connect_string = connect_string + "Server=" + db_host + ";"
    connect_string = connect_string + "Database=" + db_name + ";"
    if (db_trusted == 'true'):
        connect_string = connect_string + "Trusted_Connection=yes;"
    else:
        connect_string = connect_string + "Uid=" + db_user + ";"
        connect_string = connect_string + "Pwd=" + db_pwd + ";"
    return connect_string