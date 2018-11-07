from selenium import webdriver
from selenium.webdriver.support.select import Select
import os
import time
import pandas as pd
# SMTPLIB to send email
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import sys
from datetime import date, datetime
import glob



def prodCheck(log_data):
    #import time

    # Initialize WebDriver
    options = webdriver.ChromeOptions()
    # options.add_argument("--user-data-dir=/tmp/selenium/chrome")
    driver = webdriver.Chrome(options=options, executable_path="C:\\Reports\\chromedriver.exe")
    # Set URL
    #URL = 'http://localhost:5619/login.max?preprocess=true'
    URL = 'https://login.relyonanchor.com/login.max?preprocess=true'
    # URL = 'http://52.206.92.78:5619/login.max?preprocess=true'
    # Load the Login Page
    driver.get(URL)

    # Login to the application
    try:
        username = driver.find_element_by_name('UserName')
        password = driver.find_element_by_name('password')
        submit = driver.find_element_by_name('submit')
        #username.send_keys("maxuser")
        username.send_keys("tbobbili")
        #password.send_keys("Test123!#")
        password.send_keys("Anchor@1234")
        submit.click()
        print("Login to application is Successful!")
        time.sleep(5)
    except:
        print("Stingray is down")
        # Define from and to address
        msg = MIMEMultipart()
        msg["Subject"] = time.strftime("%m/%d") + " Stingray Daily Prod Check - Automated Test - Stingray Down"
        msg["From"] = "tbobbili@relyonanchor.com"
        # msg["To"] = "tbobbili@relyonanchor.com; Indrajith Mandapati <IMandapati@relyonanchor.com>; Nitish Settipalli <nsettipalli@relyonanchor.com>; Ramesh Pabbu <rpabbu@relyonanchor.com>; Surendra Pepakayala <SPepakayala@relyonanchor.com>"
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

    # Extract data
    rows_list = []
    col = ['Job_Name', 'Job_status', 'Schedule', 'Last_run_date', 'Last_run_status']
    df = pd.DataFrame(rows_list, columns=col)

    for i in range(3, 246, 5):
        task_name = (
            driver.find_element_by_xpath("//*[@id='rightColumn']/div[3]/table/tbody/tr[" + str(i) + "]/td[1]")).text
        job_status = (
            driver.find_element_by_xpath("//*[@id='rightColumn']/div[3]/table/tbody/tr[" + str(i) + "]/td[2]")).text
        schedule = Select(driver.find_element_by_xpath(
            "//*[@id='rightColumn']/div[3]/table/tbody/tr[" + str(i) + "]/td[3]/select")).first_selected_option.text
        last_run_date = (
            driver.find_element_by_xpath("//*[@id='rightColumn']/div[3]/table/tbody/tr[" + str(i) + "]/td[4]")).text
        last_run_status = (
            driver.find_element_by_xpath("//*[@id='rightColumn']/div[3]/table/tbody/tr[" + str(i + 2) + "]/td[1]")).text
        rows_list = [task_name, job_status, schedule, last_run_date, last_run_status]
        df.loc[i - 2] = rows_list

        # print (task_name , str(1), last_run_date, str(1), status, str(1), schedule, str(1), last_run_status )
    driver.quit()
    # pd.set_option('display.width', 200)
    # print (df)
    # df.to_csv("C:/Users/tbobbili/Desktop/output.csv", index=False)

    # Daily Jobs in running Status
    # print (df.loc[(df.job_status=='Enabled') & (df.schedule=='Daily') & (df.last_run_status.str.contains("Running"))])
    running_df = df.loc[
        (df.Job_status == 'Enabled') & (df.Schedule == 'Daily') & (df.Last_run_status.str.contains("Running"))]
    running_df = running_df.reset_index()
    running_df.index += 1

    # Daily jobs in Error status
    # print (df.loc[(df.job_status=='Enabled') &(df.last_run_status.str.contains("Error"))])
    error_df = df.loc[(df.Job_status == 'Enabled') & (
            df.Last_run_status.str.contains("Problem") | df.Last_run_status.str.contains("Error"))]
    error_df = error_df.reset_index()
    error_df.index += 1

    runningJobTable = '<table><thead><tr><th>Task Name</th><th>Last Run Status</th></tr></thead><tbody>'

    for index, row in running_df.iterrows():
        runningJobTable = runningJobTable + '<tr>' + '<td>' + row.Job_Name + '</td>' + '<td>' + \
                          (row.Last_run_status.split("Last Run Status:"))[1] + '</td>' + '</tr>'

    runningJobTable = runningJobTable + '</tbody></table>'

    errorJobTable = '<table><thead><tr><th>Task Name</th><th>Last Run Status</th></tr></thead><tbody>'

    for index, row in error_df.iterrows():
        errorJobTable = errorJobTable + '<tr>' + '<td>' + row.Job_Name + '</td>' + '<td>' + \
                        (row.Last_run_status.split("Last Run Status:"))[1] + '</td>' + '</tr>'

    errorJobTable = errorJobTable + '</tbody></table>'

    # Prepare Body of the Email
    # Set Body_Header
    if int(time.strftime("%H")) >= 17:
        body_GreetMsg = "Good Evening,<br>"
    elif int(time.strftime("%H")) >= 12:
        body_GreetMsg = "Good Afternoon,<br>"
    else:
        body_GreetMsg = "Good Morning,<br>"

    # Set Body_1 if no jobs are running set it to empty
    if running_df.empty:
        runningJobBody = ""
    else:
        runningJobBody = "Below jobs are currently in running status: <br><br>" + runningJobTable

    # Set Body_2 if no jobs are errored out set it to empty
    if error_df.empty:
        errorJobBody = ""
    else:
        errorJobBody = "Below jobs are currently in error status: <br><br>" + errorJobTable

    # Set Final_Body
    if len(runningJobBody) == 0 and len(errorJobBody) == 0:
        bodyEndMsg = "All jobs are completed." + "<br><br>-Thanks"
    else:
        bodyEndMsg = "<br>-Thanks"

    headContent = "<head><style> table, th, td { border: 1px solid black; border-collapse: collapse; padding: 5px; } th { font-weight: bold; text-align: left; background-color: #c8cbd1; } th, td { padding: 10px; }</style></head>"

    print(body_GreetMsg)
    print(runningJobBody)
    print(errorJobBody)
    print(bodyEndMsg)


    print ("Checking if any errors to report...")
    if error_df.empty:
        print ("No Errors to report")
        emailBody = headContent + body_GreetMsg + runningJobBody + "<br>" + errorJobBody + bodyEndMsg
        sendEmail_to_StingrayTeam(emailBody)
    else:
        if "Email sent to Sapiens" in log_data:
            print ("Error already reported")
            emailBody = headContent + body_GreetMsg + runningJobBody + "<br>" + errorJobBody + bodyEndMsg
            sendEmail_to_StingrayTeam(emailBody)
        else:
            emailBody = headContent + body_GreetMsg + runningJobBody + "<br>" + errorJobBody + "<br>Can you please look into it.<br>" + bodyEndMsg
            sendEmail_to_Sapiens(emailBody)



def sendEmail_to_StingrayTeam(emailBody):
    #emailBody = headContent + body_GreetMsg + runningJobBody + "<br>" + errorJobBody + "<br>Can you please look into it." + bodyEndMsg
    # Define from and to address
    msg = MIMEMultipart()
    msg["Subject"] = time.strftime("%m/%d") + " Stingray Daily Prod Check - Automated Test - Status"
    msg["From"] = "StingrayDev@relyonanchor.com"
    to_address = ['tbobbili@relyonanchor.com', 'imandapati@relyonanchor.com', 'nsettipalli@relyonanchor.com', 'rpabbu@relyonanchor.com']
    #to_address = ['tbobbili@relyonanchor.com']
    msg["To"] = ", ".join(to_address)
    # msg2["CC"] = "imandapati@relyonanchor.com"
    # Add Body
    msg.attach(MIMEText(emailBody , 'html', 'utf-8'))
    # Send Email
    s2 = smtplib.SMTP("relyonanchor-com.mail.protection.outlook.com")
    s2.sendmail("StingrayDev@relyonanchor.com", to_address, msg.as_string())
    s2.quit()
    print ("Email sent to Stingray Team")

def sendEmail_to_Sapiens(emailBody):
    #emailBody = headContent + body_GreetMsg + runningJobBody + "<br>" + errorJobBody + "<br>Can you please look into it." + bodyEndMsg
    # Define from and to address
    msg = MIMEMultipart()
    msg["Subject"] = time.strftime("%m/%d") + " Stingray Daily Prod Check - Automated Test - Error"
    msg["From"] = "StingrayDev@relyonanchor.com"
    to_address = ['tbobbili@relyonanchor.com', 'imandapati@relyonanchor.com', 'nsettipalli@relyonanchor.com', 'rpabbu@relyonanchor.com']
    #to_address = ['tbobbili@relyonanchor.com']
    msg["To"] = "tbobbili@relyonanchor.com;"
    msg["CC"] = ", ".join(to_address)
    # Add Body
    msg.attach(MIMEText(emailBody , 'html', 'utf-8'))
    # Send Email
    s2 = smtplib.SMTP("relyonanchor-com.mail.protection.outlook.com")
    s2.sendmail("StingrayDev@relyonanchor.com", to_address, msg.as_string())
    s2.quit()
    print ("Email sent to Sapiens")


log_path = 'C:/Reports/Prod Check Logs/'
todayDate = date.today().strftime("%m%d%Y")
os.chdir(log_path)
logfile_searchresult = glob.glob(todayDate+'*.log')
#print (logfile_result)
if len(logfile_searchresult) != 0:
    log_file = open(log_path + logfile_searchresult[0], "a+")

    with open(log_path + logfile_searchresult[0], 'r') as f:
        data = f.read()
        #print(data)
        latest_statuscheck = data.split("------------------------------")[-1]
        #print (data.split("------------------------------")[-1])
    sys.stdout = log_file
    print("------------------------------")
    print(datetime.now().strftime("%H:%M"))
    if 'All jobs are completed' in latest_statuscheck:
        print("All jobs are completed. Exiting Prod Check")
    else:
        prodCheck(data)

else:
    log_file = open(log_path + todayDate + " - Daily Prod Check.log", "w")
    #sys.stdout = log_file
    sys.stdout = log_file
    print("------------------------------")
    print(datetime.now().strftime("%H:%M"))
    data = ""
    prodCheck(data)





log_file.close()
