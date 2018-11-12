import requests
from datetime import timedelta, time, datetime
import time as t
from lxml import etree
import pyodbc
import util
from configparser import ConfigParser

config = ConfigParser()
config.read('./config.ini')

from_address = config.get('email', 'from_address')
to_list = [] if config.get('email', 'to_list').split(',') == [''] else config.get('email', 'to_list').split(',')
cc_list = [] if config.get('email', 'cc_list').split(',') == [''] else config.get('email', 'cc_list').split(',')

def is_env_up(env):

    url = config.get(env, 'url')
    username = config.get(env, 'user')
    password = config.get(env, 'pwd')

    ts = int(t.time() * 1000)
    try:
        response = requests.post(url + "?t=" + str(ts), data={'SubmitAct':'submit', 'Username':username, 'password':password})
        return True if ((response.status_code == 200)
                    and (etree.fromstring(response.content).xpath('//status')[0].text == 'success')) else False
    except Exception as e:
        return False

def is_dwh_current():

    connect_string = util.get_db_connection_string(config.get('database', 'db_host'), config.get('database', 'db_name'),
                                                   config.get('database', 'trusted'), config.get('database', 'db_user'),
                                                   config.get('database', 'db_pwd'))
    try:
        conn = pyodbc.connect(connect_string)
        cursor = conn.cursor()
        query = "select top 1 cast(DTS_LastRunDate as datetime) from DTS_ScheduleTask order by 1 desc"
        cursor.execute(query)
        result = cursor.fetchall()
        cursor.close()
        dwh_last_entry_date = result[0][0]

        if dwh_last_entry_date < datetime.combine(datetime.now().date() + timedelta(days=-1), time(13, 00)):
            return False
    except Exception as e:
        return False

    return True

def send_email(env_status):

    header = "<head><style> table, th, td { border: 1px solid black; border-collapse: collapse; padding: 5px; } " \
              "th { font-weight: bold; text-align: left; background-color: #c8cbd1; } " \
              "th, td { padding: 10px; }</style></head>"

    table = '<table><thead><tr><th>Environment</th><th>Status</th></tr></thead><tbody>'
    for env, status in env_status.items():
        table = table + '<tr><td>' + env + '</td><td>'
        table = table + ('<b style="color:Red">' + status + '</b>') if status in ('Down', 'Old Data') else table + status
        table = table + '</td></tr>'
    table = table + '</table>'

    msg = "Good Morning,<br><br>"
    msg = msg + "Below is the status of all Stingray environments: <br><br>" + table
    msg = msg + "<br><br>Thank you,<br>Stingray IT"
    email = header + msg

    util.send_mail(config.get('email', 'smtp_host'), config.get('email', 'smtp_user'), config.get('email', 'smtp_pwd'),
        from_address, to_list, cc_list,
        "Stingray Daily Environment Check for " + t.strftime("%m/%d"),
        email,
        [], 'html'
    )

if __name__ == "__main__":

    env_list = ['QA', 'UAT', 'STAGING']
    env_status = {}

    while (len(env_status) == 0 or 'Down' in env_status.values() or 'Old Data' in env_status.values()):
        for env in env_list:
            env_status[env] = 'Up & Running' if is_env_up(env) else 'Down'

        env_status['DWH'] = 'Latest Data' if is_dwh_current() else 'Old Data'
        send_email(env_status)
        t.sleep(int(config.get('report', 'retry')))
