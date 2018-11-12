import requests
from os.path import join
from datetime import timedelta, time, datetime, date
import time as t
import util
from configparser import ConfigParser
from lxml import etree
from bs4 import BeautifulSoup

config = ConfigParser()
config.read('./config.ini')

smtp_host = config.get('email', 'smtp_host')
smtp_user = config.get('email', 'smtp_user')
smtp_pwd = config.get('email', 'smtp_pwd')
from_address = config.get('email', 'from_address')
to_list = [] if config.get('email', 'to_list').split(',') == [''] else config.get('email', 'to_list').split(',')
cc_list = [] if config.get('email', 'cc_list').split(',') == [''] else config.get('email', 'cc_list').split(',')
dev_list = [] if config.get('email', 'dev_list').split(',') == [''] else config.get('email', 'dev_list').split(',')

output_folder = config.get('folders', 'output_folder')
log_folder = config.get('folders', 'log_folder')
log_file = join(log_folder, "daily_prod_check_" + date.today().strftime("%m%d%Y") + ".log")

def login(env):

    url = config.get(env, 'url')
    username = config.get(env, 'user')
    password = config.get(env, 'pwd')

    ts = int(t.time() * 1000)
    try:
        response = requests.post(url + "/login.max?t=" + str(ts), data={'SubmitAct':'submit', 'Username':username, 'password':password})
        if ((response.status_code != 200) and
                (etree.fromstring(response.content).xpath('//status')[0].text != 'success')):
            raise Exception("Login failed")

        token = etree.fromstring(response.content).xpath('//id')[0].text
        return token

    except Exception as e:
        util.log(log_file, repr(e), False)
        raise Exception("Login failed")

def get_tasks(env, token):

    url = config.get(env, 'url')
    ts = int(t.time() * 1000)

    try:
        response = requests.post(url + "/DTS_Scheduler.max?t=" + str(ts), data={'id': token, 'preprocess': 'true'})
        if ((response.status_code != 200) and
                (etree.fromstring(response.content).xpath('//status')[0].text != 'success')):
            raise Exception("Failed to get tasks")

        root = etree.fromstring(response.content)
        html = root.xpath('//html')[0]
        soup = BeautifulSoup(html.text, 'html.parser')
        table = soup.find_all('table')

        task_names = table[0].find_all(attrs={"rowspan": "2"})
        task_enabled = table[0].find_all(attrs={"class": "enablelink"})
        task_status = table[0].find_all(attrs={"colspan": "5", "style": ""})

        tasks = {name.string: status.text for name, enable, status in
                 zip(task_names, task_enabled, task_status) if enable.string == 'Enabled'}

        running_tasks = {k: v for k, v in tasks.items() if 'Running' in v}
        failed_tasks = {k: v for k, v in tasks.items() if 'Problem' in v}

        return (running_tasks, failed_tasks)

    except Exception as e:
        util.log(log_file, repr(e), False)
        raise Exception("Failed to get tasks")

def send_email(running_tasks, failed_tasks):

    header = "<head><style> table, th, td { border: 1px solid black; border-collapse: collapse; padding: 5px; } " \
              "th { font-weight: bold; text-align: left; background-color: #c8cbd1; } " \
              "th, td { padding: 10px; }</style></head>"

    running_table = '<table><thead><tr><th>Task</th><th>Status</th></tr></thead><tbody>'
    for task, status in running_tasks.items():
        running_table = running_table + '<tr><td>' + task + '</td><td>' + status + '</td></tr>'
    running_table = running_table + '</table>'

    failed_table = '<table><thead><tr><th>Task</th><th>Status</th></tr></thead><tbody>'
    for task, status in failed_tasks.items():
        failed_table = failed_table + '<tr><td>' + task + '</td><td>' + '<b style="color:Red">' + status + '</b></td></tr>'
    failed_table = failed_table + '</table>'

    msg = "Good Morning,<br><br>"
    msg = msg + "Below is the list of tasks still running in Stingray production:<br><br>" + running_table
    msg = msg + "<br><br>Below is the list of tasks that failed in Stingray production:<br><br>" + failed_table
    msg = msg + "<br><br>Thank you,<br>Stingray IT"
    email = header + msg

    util.send_mail(config.get('email', 'smtp_host'), config.get('email', 'smtp_user'), config.get('email', 'smtp_pwd'),
        from_address, to_list, cc_list,
        "Stingray Daily Environment Check for " + t.strftime("%m/%d"),
        email,
        [], 'html'
    )

def send_failure_email(file):
    util.send_mail(smtp_host, smtp_user, smtp_pwd,
        from_address, dev_list, [],
        "Production check failed",
        "Good morning,\n\n" \
            + "Note: Production check failed. Please check the attached log file and take action if needed " \
            + "The process will check in 60 mins and send a status if successful \n\n" \
            + "Thank you, Stingray IT",
        [file]
    )

if __name__ == "__main__":

    util.log(log_file, "Process started", False)
    done = False
    while not done:
        try:
            token = login('PROD')
            running_tasks, failed_tasks = get_tasks('PROD', token)
            send_email(running_tasks, failed_tasks)
            util.log(log_file, "Sent status email", False)
            done = (len(running_tasks) == 0 and len(failed_tasks) == 0)
        except Exception as e:
            util.log(log_file, repr(e), False)
            send_failure_email(log_file)
            util.log(log_file, "Sent failure email", False)
            done = False

        t.sleep(int(config.get('report', 'retry')))
    util.log(log_file, "Process ended", False)
