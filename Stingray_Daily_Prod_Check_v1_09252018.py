
from selenium import webdriver
from selenium.webdriver.support.select import Select
import os
import time
import pandas as pd
# SMTPLIB to send email
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

#Initialize WebDriver
options = webdriver.ChromeOptions()
#options.add_argument("--user-data-dir=/tmp/selenium/chrome")
driver = webdriver.Chrome(options=options, executable_path="C:\\Reports\\chromedriver.exe")
# Set URL
#URL = 'http://localhost:5619/login.max?preprocess=true'
#URL = 'https://login.relyonanchor.com/login.max?preprocess=true'
URL = 'http://52.206.92.78:5619/login.max?preprocess=true'
#Load the Login Page
driver.get(URL)

#Login to the application
try:
    username = driver.find_element_by_name('UserName')
    password = driver.find_element_by_name('password')
    submit = driver.find_element_by_name('submit')
    username.send_keys("maxuser")
    #username.send_keys("tbobbili")
    password.send_keys("Test123!#")
    #password.send_keys("Anchor@1234")
    submit.click()
    print ("Login to application is Successful!")
    time.sleep(5)
except:
    print("Stingray is down")
    # Define from and to address
    msg = MIMEMultipart()
    msg["Subject"] = "Stingray Production is Down - Automated Test"
    msg["From"] = "tbobbili@relyonanchor.com"
    #msg["To"] = "tbobbili@relyonanchor.com; Indrajith Mandapati <IMandapati@relyonanchor.com>; Nitish Settipalli <nsettipalli@relyonanchor.com>; Ramesh Pabbu <rpabbu@relyonanchor.com>; Surendra Pepakayala <SPepakayala@relyonanchor.com>"
    msg["To"] = "tbobbili@relyonanchor.com"
    # Add Body
    msg.attach(MIMEText(
        "Good Morning,\n\nStingray production is not accessible. Please look into it asap.\n\nThanks.",
        'plain', 'utf-8'))
    # Send Email
    s = smtplib.SMTP("relyonanchor-com.mail.protection.outlook.com")
    s.sendmail("tbobbili@relyonanchor.com", "tbobbili@relyonanchor.com", msg.as_string())
    s.quit()
    exit()

# Open SST Scheduler Page
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
wait = WebDriverWait(driver, 10)
click_system_admin_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "System Administration")))
driver.execute_script("Communication.LinkRequest('DTS_Scheduler.max');")
time.sleep(3)

#Extract data
rows_list = []
col = ['Job_Name', 'Job_status', 'Schedule','Last_run_date','Last_run_status']
df = pd.DataFrame(rows_list,columns=col)

for i in range(3,246,5):
    task_name = (driver.find_element_by_xpath("//*[@id='rightColumn']/div[3]/table/tbody/tr["+str(i)+"]/td[1]")).text
    job_status = (driver.find_element_by_xpath("//*[@id='rightColumn']/div[3]/table/tbody/tr["+str(i)+"]/td[2]")).text
    schedule = Select(driver.find_element_by_xpath("//*[@id='rightColumn']/div[3]/table/tbody/tr["+str(i)+"]/td[3]/select")).first_selected_option.text
    last_run_date = (driver.find_element_by_xpath("//*[@id='rightColumn']/div[3]/table/tbody/tr["+str(i)+"]/td[4]")).text
    last_run_status = (driver.find_element_by_xpath("//*[@id='rightColumn']/div[3]/table/tbody/tr["+str(i+2)+"]/td[1]")).text
    rows_list = [task_name ,  job_status, schedule, last_run_date, last_run_status]
    df.loc[i-2] = rows_list
    #print (task_name , str(1), last_run_date, str(1), status, str(1), schedule, str(1), last_run_status )

#pd.set_option('display.width', 200)
#print (df)
#df.to_csv("C:/Users/tbobbili/Desktop/output.csv", index=False)


# Daily Jobs in running Status
#print (df.loc[(df.job_status=='Enabled') & (df.schedule=='Daily') & (df.last_run_status.str.contains("Running"))])
running_df = df.loc[(df.Job_status=='Enabled') & (df.Schedule=='Daily') & (df.Last_run_status.str.contains("Running"))]
running_df = running_df.reset_index()
running_df.index += 1
#Daily jobs in Error status
#print (df.loc[(df.job_status=='Enabled') &(df.last_run_status.str.contains("Error"))])
error_df = df.loc[(df.Job_status=='Enabled') &(df.Last_run_status.str.contains("Problem") | df.Last_run_status.str.contains("Error"))]
error_df = error_df.reset_index()
error_df.index += 1

#Prepare Body of the Email
from tabulate import tabulate
# Set Body_Header
import time
if int(time.strftime("%H")) >= 17:
    body_header = "Good Evening,\n"
elif int(time.strftime("%H")) >= 12:
    body_header = "Good Afternoon,\n"
else:
    body_header = "Good Morning,\n"

# Set Body_1 if no jobs are running set it to empty
body_1 = "\nBelow jobs are currently in running status: \n" #+ str(running_df.Job_Name.tolist())  + "\n"
if len(running_df.Job_Name.tolist()) != 0:
    body_1 = body_1 + tabulate(running_df[['Job_Name']],  tablefmt='plain') + "\n"
else:
    body_1 = ""

# Set Body_2 if no jobs are errored out set it to empty
if error_df.empty:
    body_2 = ""
else:
    body_2 = "Below jobs are currently in error status:\n\n" + tabulate(error_df[['Job_Name', 'Last_run_status']], tablefmt='plain')


# Set Final_Body
if len(body_1) == 0 and len(body_2) == 0:
    final_body = "All jobs are completed" + "\n\n-Thanks"
else:
    final_body = "\n-Thanks"

print (body_1)
print (body_2)
print (final_body)


# Define from and to address
msg2 = MIMEMultipart()
msg2["Subject"] = "Stingray Daily Prod Check - Automated Test"
msg2["From"] = "tbobbili@relyonanchor.com"
#msg2["To"] = "tbobbili@relyonanchor.com"
msg2["To"] = "tbobbili@relyonanchor.com; Indrajith Mandapati <IMandapati@relyonanchor.com>; Nitish Settipalli <nsettipalli@relyonanchor.com>; Ramesh Pabbu <rpabbu@relyonanchor.com>; Surendra Pepakayala <SPepakayala@relyonanchor.com>; Shimpy <shimpy@relyonanchor.com>"
# Add Body
msg2.attach(MIMEText(  body_header + body_1 +"\n"+ body_2 + "\n" +final_body ,    'plain', 'utf-8'))
# Send Email
s2 = smtplib.SMTP("relyonanchor-com.mail.protection.outlook.com")
s2.sendmail("tbobbili@relyonanchor.com", "tbobbili@relyonanchor.com; Indrajith Mandapati <IMandapati@relyonanchor.com>; Nitish Settipalli <nsettipalli@relyonanchor.com>; Ramesh Pabbu <rpabbu@relyonanchor.com>; Surendra Pepakayala <SPepakayala@relyonanchor.com>; Shimpy <shimpy@relyonanchor.com>", msg2.as_string())
#s2.sendmail("tbobbili@relyonanchor.com", "tbobbili@relyonanchor.com", msg2.as_string())
s2.quit()
#time.sleep(5)
print ("Daily Prod Check is complete!")
driver.quit()