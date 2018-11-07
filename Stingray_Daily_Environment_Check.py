from selenium import webdriver
import time

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import pyodbc

# Initialize environment constants
run_mode = 'QA'  # Options: Dev / QA / Prod

if run_mode == 'Dev':
    email_to_list = ['imandapati@relyonanchor.com']
    include_support_team = 0
elif run_mode == 'QA':
    email_to_list = ['nsettipalli@relyonanchor.com', 'rpabbu@relyonanchor.com', 'imandapati@relyonanchor.com', 'tbobbili@relyonanchor.com']
    include_support_team = 0
elif run_mode == 'Prod':
    email_to_list = ['ccarpenter@Relyonanchor.com', 'nzachariah@relyonanchor.com', 'jvanko@relyonanchor.com', 'rmydam@relyonanchor.com', 'spepakayala@relyonanchor.com',  'tbobbili@relyonanchor.com', 'imandapati@relyonanchor.com', 'nsettipalli@relyonanchor.com', 'rpabbu@relyonanchor.com']
    include_support_team = 1 if time.localtime().tm_hour < 8 else 0
else:
    print('Environment setup is wrong')
    exit()

# Initialize WebDriver
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(options=options, executable_path="C:/Reports/bin/chromedriver.exe")

# QA Check
try:
    URL = 'http://52.207.35.39:5619/login.max?preprocess=true'
    driver.get(URL)
    username = driver.find_element_by_name('UserName')
    password = driver.find_element_by_name('password')
    submit = driver.find_element_by_name('submit')
    username.send_keys("maxuser")
    password.send_keys("Test123!#")
    submit.click()
    print("Login to QA application is Successful!")
    qa_status = 'Up & Running'
    time.sleep(5)
except:
    print("QA is down")
    qa_status = '<b style="color:Red">Down</b>'

# UAT Check
try:
    URL = 'http://52.206.92.78:5619/login.max?preprocess=true'
    driver.get(URL)
    username = driver.find_element_by_name('UserName')
    password = driver.find_element_by_name('password')
    submit = driver.find_element_by_name('submit')
    username.send_keys("maxuser")
    password.send_keys("Test123!#")
    submit.click()
    print("Login to UAT application is Successful!")
    uat_status = 'Up & Running'
    time.sleep(5)
except:
    print("UAT is down")
    uat_status = '<b style="color:Red">Down</b>'

# Staging Check
try:
    URL = 'http://staging.relyonanchor.com:5619/login.max?preprocess=true'
    driver.get(URL)
    username = driver.find_element_by_name('UserName')
    password = driver.find_element_by_name('password')
    submit = driver.find_element_by_name('submit')
    username.send_keys("maxuser")
    password.send_keys("Test123!#")
    submit.click()
    print("Login to Staging application is Successful!")
    staging_status = 'Up & Running'
    time.sleep(5)
except:
    print("Stingray is down")
    staging_status = '<b style="color:Red">Down</b>'

driver.quit()

# DWH Check
try:
    # Connection to Database
    dbConn = pyodbc.connect('Driver={SQL Server Native Client 11.0};Server=dwh.relyonanchor.com;Database=AIH_Insurance;Trusted_Connection=yes;')
    cursor = dbConn.cursor()

    # Query for daily balance forward/ renewal notification report for Underwriting taken from SRD-861
    # Begin Query and execution
    query = "SELECT CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END AS 'QuotesExist' FROM PolicyQuote WHERE CAST(PQ_EntryDate AS DATE) = CAST(GETDATE()-1 AS DATE)"

    # Execute the query
    cursor.execute(query)
    # Save the results in result_set2 variable
    result_set = cursor.fetchone()

    if result_set.QuotesExist == 1:
        dwh_status = 'Up to date'
    else:
        dwh_status = '<b style="color:Red">Old Data</b>'

except:
    print("DWH is down")
    dwh_status = 'Down'

env_table = '<table><thead><tr><th>Environment</th><th>Status</th></tr></thead><tbody>'
env_table = env_table + '<tr><td>QA</td><td>' + qa_status + '</td></tr>'
env_table = env_table + '<tr><td>UAT</td><td>' + uat_status + '</td></tr>'
env_table = env_table + '<tr><td>Staging</td><td>' + staging_status + '</td></tr>'
env_table = env_table + '<tr><td>DWH</td><td>' + dwh_status + '</td></tr>'
env_table = env_table + '</table>'

headContent = "<head><style> table, th, td { border: 1px solid black; border-collapse: collapse; padding: 5px; } th { font-weight: bold; text-align: left; background-color: #c8cbd1; } th, td { padding: 10px; }</style></head>"

greet_msg = "Good Morning,<br><br>"
env_status = "Below is the status of all Stingray environments: <br><br>" + env_table
thanks_msg = "<br><br>Thank you,<br>Stingray IT"
emailContent = headContent + greet_msg + env_status + thanks_msg

# Define from and to address
msg2 = MIMEMultipart()
if run_mode == 'Prod':
    subject = time.strftime("%m/%d") + " Stingray Daily Environment Check"
else:
    subject = time.strftime("%m/%d") + " Stingray Daily Environment Check " + '- Automated Test'

msg2["Subject"] = subject
msg2["From"] = "StingrayDev@relyonanchor.com"

if dwh_status != 'Up to date' and include_support_team == 1:
    email_to_list.append('helpdesk@relyonanchor.com')

msg2["To"] = ", ".join(email_to_list)

# Add Body
msg2.attach(MIMEText(emailContent, 'html', 'utf-8'))
# Send Email
s2 = smtplib.SMTP("relyonanchor-com.mail.protection.outlook.com")
s2.sendmail("StingrayDev@relyonanchor.com", email_to_list, msg2.as_string())
s2.quit()

print("Daily Environment Check is completed!")