import requests
from os.path import join
from datetime import timedelta, time, datetime, date
import time as t
from lxml import etree
import pyodbc
import util
from configparser import ConfigParser

config = ConfigParser()
config.read('./config.ini')

from_address = config.get('email', 'from_address')
to_list = [] if config.get('email', 'to_list').split(',') == [''] else config.get('email', 'to_list').split(',')
admin_list = [] if config.get('email', 'admin_list').split(',') == [''] else config.get('email', 'admin_list').split(',')
dev_list = [] if config.get('email', 'dev_list').split(',') == [''] else config.get('email', 'dev_list').split(',')

output_folder = config.get('folders', 'output_folder')
log_folder = config.get('folders', 'log_folder')
log_file = join(log_folder, "daily_env_" + date.today().strftime("%m%d%Y") + ".log")

def is_env_up(env, env_status):

    url = config.get(env, 'url')
    username = config.get(env, 'user')
    password = config.get(env, 'pwd')

    ts = int(t.time() * 1000)
    try:
        response = requests.post(url + "?t=" + str(ts), data={'SubmitAct':'submit', 'Username':username, 'password':password})
        if ((response.status_code != 200)
                    or (etree.fromstring(response.content).xpath('//status')[0].text != 'success')):
            env_status[env] = 'Down'
            return env_status

        env_status[env]  = 'Up & Running'
        return env_status

    except Exception as e:
        util.log(log_file, repr(e), False)
        env_status[env] = 'Down'
        return env_status

def is_dwh_current(env, env_status):

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
            env_status[env] = 'Old Data'
            return env_status

        env_status[env] = 'Up to date'
        return env_status

    except Exception as e:
        util.log(log_file, repr(e), False)
        env_status[env] = 'Old Data'
        return env_status

def send_email(env_status, file=[], admin_list=[]):

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
        from_address, to_list, admin_list,
        "Stingray Daily Environment Check for " + t.strftime("%m/%d"),
        email, file, 'html'
    )
    return True

def send_failure_email(file=[]):
    util.send_mail(config.get('email', 'smtp_host'), config.get('email', 'smtp_user'), config.get('email', 'smtp_pwd'),
        from_address, dev_list, [],
        "Stingray Daily Environment Check failed",
        "Good morning,\n\n" \
            + "Note: Stingray Daily environment check failed. Please check the attached log file and take action if needed " \
            + "The process will try in 60 mins and send a status if successful \n\n" \
            + "Thank you, Stingray IT",
        file
    )

if __name__ == "__main__":

    util.log(log_file, "Process started", False)
    admin_email = False
    env_list = [] if config.get('report', 'environments').split(',') == [''] else config.get('report', 'environments').split(',')
    env_status = {}

    while (len(env_status) == 0 or 'Down' in env_status.values() or 'Old Data' in env_status.values()):
        try:
            for env in env_list:
                env_status = is_env_up(env.strip(), env_status)
            env_status = is_dwh_current('DWH', env_status)

            if (not admin_email and datetime.now().hour >= 8):
                admin_email = send_email(env_status, file=[log_file], admin_list=admin_list)
            else:
                send_email(env_status)

        except Exception as e:
            util.log(log_file, repr(e), False)
            send_failure_email([log_file])

        t.sleep(int(config.get('report', 'retry')))

    util.log(log_file, "Process ended", False)