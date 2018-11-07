import pyodbc
import xlwt
from os.path import join
from xlwt import *
from datetime import date, timedelta

def run_report(inputdate):

    # Connection to Dataware House
    cnxn = pyodbc.connect('Driver={SQL Server Native Client 11.0};Server=dwh.relyonanchor.com;Database=AIH_Insurance;Trusted_Connection=yes;')
    cursor = cnxn.cursor()

    # Query for daily balance forward/ renewal notification report for Underwriting taken from SRD-861
    # Begin Query and execution
    q1 = """SELECT A.PD_PolicyCode as 'Policy Code'	, PC.PC_FirstName as 'First Name'	, PC.PC_LastName as 'Last Name'	, SCT.sCT_Code as 'Agency Code'	, SCT.sCT_Name as 'Agency Name' \
             , cast(A.PD_InceptionDate as date) as 'Inception Date'	, cast(A.PD_ExpirationDate as date) as 'Expiration Date' , cast(PP.PPL_NumOfPayments as varchar)  + ' Pay - '+ PP.PPL_Display as 'Pay Plan' \
             , ROUND(SUM(PAC.pACC_Debit - PAC.pACC_Credit),2) as 'Amount Outstanding' , A.PD_Total_Premium as 'Policy Premium'	, SCT.sCT_EMail as 'Agent Email'	, SCT.sCT_Phone as 'Agent Phone Number' \
             FROM PolicyData A  \
             LEFT JOIN PolicyAccounting PAC ON A.PD_PolicyCode = PAC.PD_PolicyCode \
             LEFT JOIN Policy_Contacts PC ON A.PC_ID = PC.PC_ID \
             LEFT JOIN PayPlans AS PP with (NOLOCK) ON PP.PPL_ID = A.PPL_ID \
             LEFT JOIN System_CompanyTree SCT ON A.PD_Agency = SCT.SCT_ID \
             WHERE A.sPTRN_ID IN (3,4,5,8) and A.PD_CurrentRecord = 1 and A.PD_ExpirationDate < GETDATE() + 100 AND A.PD_ExpirationDate >= GETDATE() \
             GROUP BY A.PD_PolicyCode, PC.PC_FirstName, PC.PC_LastName, SCT.sCT_Code, SCT.sCT_Name, cast(A.PD_InceptionDate as date), cast(A.PD_ExpirationDate as date), cast(PP.PPL_NumOfPayments as varchar)  + ' Pay - '+ PP.PPL_Display	, A.PD_Total_Premium, SCT.sCT_EMail	, SCT.sCT_Phone \
             HAVING ROUND(SUM(PAC.pACC_Debit - PAC.pACC_Credit),2) > 0; """
    # Execute the query
    cursor.execute(q1)
    # Save the results in result_set2 variable
    result_set2 = cursor.fetchall()

    # Create a Excel Workbook
    wb = Workbook()
    # Add a sheet SIS NB Inspection
    ws0 = wb.add_sheet('Sheet 1')
    # Set Style_string for headers & rows
    style_string = "font: bold off, height 220, name Calibri"
    style = xlwt.easyxf(style_string)
    style_string1 = "font: bold on, height 220, name Calibri; pattern: pattern solid, fore_colour gray25; borders: top_color black, bottom_color black, right_color black, left_color black,\
                                  left thin, right thin, top thin, bottom thin;"
    style1 = xlwt.easyxf(style_string1)
    # Add headers to the excel workbook
    ws0.write(0, 0, 'Policy Code', style1)
    ws0.write(0, 1, 'First Name', style1)
    ws0.write(0, 2, 'Last Name', style1)
    ws0.write(0, 3, 'Agency Code', style1)
    ws0.write(0, 4, 'Agency Name', style1)
    ws0.write(0, 5, 'Inception Date', style1)
    ws0.write(0, 6, 'Expiration Date', style1)
    ws0.write(0, 7, 'Pay Plan', style1)
    ws0.write(0, 8, 'Amount Outstanding', style1)
    ws0.write(0, 9, 'Policy Premium', style1)
    ws0.write(0, 10, 'Agent Email', style1)
    ws0.write(0, 11, 'Agent Phone Number', style1)

    row_number = 1
    # Paste the results in the work book
    for row in result_set2:
        column_num = 0
        for item in row:  # i.e. for each field in that row
            ws0.write(row_number, column_num, str(item),
                      style)  # ,wb.get_sheet(0).cell(0,0).style)  #write excel cell from the cursor at row 1
            column_num = column_num + 1  # increment the column to get the next field

        row_number = row_number + 1
    # Set column width for the sheet
    for i in range(12):
        if i == 1:
            ws0.col(i).width = 256 * 20
        else:
            ws0.col(i).width = 256 * 15
        i = i + 1


    #Path where the reports are stored
    test_files = "C:/Reports/Daily/Stingray/Daily Balance forward - renewal notf report/"
    #Set filenames and subject name for email
    a = inputdate - timedelta(1)
    filename = a.strftime("%m%d%Y")
    subjectname = a.strftime("%m/%d/%Y")

    fname = 'Stingray Daily Balance forward Renewal Notification report - ' + filename + '.xls'
    # Save the excel file
    wb.save(join(test_files, fname))

    # SMTPLIB Code to send email
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email import encoders
    # Define from and to address
    msg = MIMEMultipart()
    msg["Subject"] = "Stingray Daily Balance forward Renewal Notification report - " + subjectname
    msg["From"] = "StingrayDev@relyonanchor.com"
    to_address = ['stgdlybalfwd@relyonanchor.com','tbobbili@relyonanchor.com', 'shimpy@relyonanchor.com']
    msg["To"] = "stgdlybalfwd@relyonanchor.com"
    msg["CC"] = "tbobbili@relyonanchor.com; Shimpy <shimpy@relyonanchor.com>"
    # Add Body
    msg.attach(MIMEText(
        "Good morning,\n\nPlease find attached daily balance forward / renewal notification report from Stingray for " + subjectname + ".\n\nThanks.",
        'plain', 'utf-8'))
    # Add attachment
    att1 = MIMEBase('application', "octet-stream")
    att1.set_payload(open(test_files + fname, 'rb').read())
    encoders.encode_base64(att1)
    att1.add_header('Content-Disposition', 'attachment; filename= ' + fname + '')
    msg.attach(att1)
    # Send Email
    s = smtplib.SMTP("relyonanchor-com.mail.protection.outlook.com")
    s.sendmail("StingrayDev@relyonanchor.com", to_address, msg.as_string())
    s.quit()

    """
    #Below Code is for Outlook
    # Code to Send email from Outlook with File Attachments
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = "Stingray Daily Balance forward Renewal Notification report - " + subjectname
    newMail.Body = "Good morning,\n\nPlease find attached daily balance forward / renewal notification report from Stingray for " + subjectname + ".\n\nThanks,\nTeja"
    # newMail.To = "tbobbili@relyonanchor.com"
    newMail.To = "dailystingraybalforwardreport@relyonanchor.com"
    newMail.CC = "Caity Carpenter <ccarpenter@Relyonanchor.com>"
    # newMail.BCC = "tbobbili@relyonanchor.com"
    attachment1 = ("C:/Users/tbobbili/Desktop/Task/Stingray Daily Bal fwd Renewal Notf report/Stingray Daily Balance forward Renewal Notification report - " + filename + ".xls")
    # attachment2 = ("C:/Users/tbobbili/Desktop/DailyActivity - 12042017.xls")
    # attachment2 = "Path to attachment no. 2"
    newMail.Attachments.Add(attachment1)
    # newMail.Attachments.Add(attachment2)
    newMail.display()
    # newMail.Send()
    print("Email generated successfully!!")
    """

def send_email_DWH_notCurrent():
    msg = MIMEMultipart()
    msg["Subject"] = "DWH is not current to send Daily Balance Forward Report"
    msg["From"] = "StingrayDev@relyonanchor.com"
    to_address = ['tbobbili@relyonanchor.com','shimpy@relyonanchor.com', 'ccarpenter@Relyonanchor.com']
    msg["To"] = "tbobbili@relyonanchor.com; Shimpy <shimpy@relyonanchor.com>; Caity Carpenter <ccarpenter@Relyonanchor.com>"
    # Add Body
    msg.attach(MIMEText("Good morning,\n\nNote: DWH is not current with latest data to send daily balance forward report. Please send the report manually once refresh is complete.\n\nThank you, Stingray IT",'plain', 'utf-8'))
    # Send Email
    s = smtplib.SMTP("relyonanchor-com.mail.protection.outlook.com")
    s.sendmail("StingrayDev@relyonanchor.com", to_address, msg.as_string())
    s.quit()

def send_email_DWH_unabletoconnect():
    msg = MIMEMultipart()
    msg["Subject"] = "Unable to connect to DWH"
    msg["From"] = "StingrayDev@relyonanchor.com"
    to_address = ['tbobbili@relyonanchor.com','shimpy@relyonanchor.com', 'ccarpenter@Relyonanchor.com']
    msg["To"] = "tbobbili@relyonanchor.com; Shimpy <shimpy@relyonanchor.com>; Caity Carpenter <ccarpenter@Relyonanchor.com>"
    # Add Body
    msg.attach(MIMEText("Good morning,\n\nNote: The python script is unable to connect to DWH. Please check and send the report manually.\n\nThank you, Stingray IT",'plain', 'utf-8'))
    # Send Email
    s = smtplib.SMTP("relyonanchor-com.mail.protection.outlook.com")
    s.sendmail("StingrayDev@relyonanchor.com", to_address, msg.as_string())
    s.quit()

def check_DWH_isCurrent(inputdate):
    try:
        #Connect to DWH
        cnxn = pyodbc.connect('Driver={SQL Server Native Client 11.0};Server=dwh.relyonanchor.com;Database=AIH_Insurance;Trusted_Connection=yes;')
    except:
        send_email_DWH_unabletoconnect()

    # Define a cursor
    cursor = cnxn.cursor()
    # Query to Check Data warehouse in latest
    q = "select top 1 cast(PD_EntryDate as date) from PolicyData order by 1 desc"
    # Execute query
    cursor.execute(q)
    # Fetch Results
    result = cursor.fetchall()
    DWH_LastEntryDate = result[0][0]
    if DWH_LastEntryDate == inputdate + timedelta(days=-1):
        print('DWH is up to date')
        return 1
    else:
        print('Send email notifying DWH is not current')
        send_email_DWH_notCurrent()
        return 0


if __name__ == "__main__":
    todaydate = date.today()

    weekno = todaydate.weekday()
    # List of Anchor hoidays. Report will not be run on below Dates
    holiday_list = ['2018-11-22', '2018-11-23']

    if todaydate.strftime('%Y-%m-%d') in holiday_list or weekno >= 5:
        print('Report not run due to company holiday')
    else:
        print('Running script')
        DWH_check = check_DWH_isCurrent(todaydate)
        if DWH_check == 1:
            run_report(todaydate)

