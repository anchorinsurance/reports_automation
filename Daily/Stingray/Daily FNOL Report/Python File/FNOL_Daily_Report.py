import xlwt
from os.path import join
from xlwt import *
from datetime import date, timedelta
import pyodbc
import pandas as pd


def run_report(inputDate):
    # Connection to FNOL Database
    dbConn = pyodbc.connect('Driver={SQL Server Native Client 11.0};Server=35.170.29.19;Database=RelyOnAnchor;uid=fnol_reader;pwd=$tingR@y;')
    cursor = dbConn.cursor()

    # Query for daily balance forward/ renewal notification report for Underwriting taken from SRD-861
    # Begin Query and execution
    query = """SELECT CD_ReferenceNumber As 'Reference Number'
      ,CONVERT(VARCHAR(10), CD_DateOfLoss, 101) As 'Date of Loss'
      ,CASE WHEN CD_HourOfLoss = '' THEN '' ELSE (CD_HourOfLoss + ':' + CD_MinuteOfLoss + ' ' + CD_MeridianOfLoss + ' ' + CD_TimeZoneOfLoss) END As 'Time of Loss'
      ,CD_PersonReporting As 'Person Reporting'
      ,CD_RelationshipToInsured As 'Relation To Insured'
      ,CD_PolicyHolderName As 'Policy Holder Name'
      ,CD_PolicyNumber As 'Policy Number'

      ,CD_PropertyAddress1 As 'Address 1'
      ,CD_PropertyAddress2 As 'Address 2'
      ,CD_PropertyCity As 'City'
      ,SS_P.sST_Description As 'State'
      ,CD_PropertyZipCode As 'Zipcode'
      ,CD_PreferredContactNumber As 'Preferred Contact Number'
      ,CD_AlternateContactNumber As 'Alternate Contact Number'
      ,CD_EmailAddress As 'Email Address'
      ,SCL.sCL_Name As 'Cause Of Loss'      

      ,CD_LossDescription As 'Loss Description'
      ,CASE WHEN CD_IsPropertyHabitable = 1 THEN 'Yes' ELSE 'No' END AS 'Is property habitable'
      ,CAST(CD.CD_CreatedDate AS VARCHAR) As 'Loss Created Date'
      ,SDC_S.sDC_Display As 'Status'
      ,ISNULL(SU.sUSR_FirstName + ' ' + SU.sUSR_LastName, '') As 'Last Modified By'
      ,CASE WHEN CD.CD_LastModifiedDate IS NULL THEN '' ELSE CAST(CD.CD_LastModifiedDate AS VARCHAR) END AS 'Last Modified Date'
  FROM RelyOnAnchor.dbo.Claims_Data AS CD WITH (NOLOCK)
  JOIN System_States AS SS_P WITH (NOLOCK) ON SS_P.sST_ID = CD.CD_PropertyStateID    
  LEFT JOIN System_CauseOfLoss AS SCL WITH (NOLOCK) ON SCL.sCL_ID = CD.sCL_CauseOfLossID
  LEFT JOIN System_CauseOfLossSubOptions AS SCLS WITH (NOLOCK) ON SCLS.sCLO_ID = CD.sCLO_CauseOfLossSubID    
  JOIN System_DataCollections AS SDC_S WITH (NOLOCK) ON SDC_S.sDC_Group = 'ClaimStatus' AND SDC_S.sDC_Code = CD.CD_StatusID  
  LEFT JOIN System_Users AS SU WITH (NOLOCK) ON SU.sUSR_ID = CD.CD_LastModifiedBy
  WHERE CAST(CD.CD_CreatedDate AS DATE) = CAST(GETDATE()-1 AS DATE) --Comment this to view all FNOL's
  ORDER BY CD_ID"""

    # Execute the query
    cursor.execute(query)
    # Save the results in result_set2 variable
    result_set = cursor.fetchall()


    df = pd.DataFrame(result_set)

    #print(df)

    # Create yesterdayDate Excel Workbook
    wb = Workbook()
    # Add yesterdayDate sheet SIS NB Inspection
    ws0 = wb.add_sheet('Data')
    # Set Style_string for headers & rows
    data_style_string = "font: bold off, height 220, name Calibri"
    data_style = xlwt.easyxf(data_style_string)
    header_style_string = "font: bold on, height 220, name Calibri; pattern: pattern solid, fore_colour gray25; borders: top_color black, bottom_color black, right_color black, left_color black, left thin, right thin, top thin, bottom thin;"

    header_style = xlwt.easyxf(header_style_string)
    # Add headers to the excel workbook
    ws0.write(0, 0, 'Reference Number', header_style)
    ws0.write(0, 1, 'Date of Loss', header_style)
    ws0.write(0, 2, 'Time of Loss', header_style)
    ws0.write(0, 3, 'Person Reporting', header_style)
    ws0.write(0, 4, 'Relation To Insured', header_style)
    ws0.write(0, 5, 'Policy Holder Name', header_style)
    ws0.write(0, 6, 'Policy Number', header_style)
    ws0.write(0, 7, 'Address 1', header_style)
    ws0.write(0, 8, 'Address 2', header_style)
    ws0.write(0, 9, 'City', header_style)
    ws0.write(0, 10, 'State', header_style)
    ws0.write(0, 11, 'Zipcode', header_style)
    ws0.write(0, 12, 'Preferred Contact Number', header_style)
    ws0.write(0, 13, 'Alternate Contact Number', header_style)
    ws0.write(0, 14, 'Email Address', header_style)
    ws0.write(0, 15, 'Cause Of Loss', header_style)
    ws0.write(0, 16, 'Loss Description', header_style)
    ws0.write(0, 17, 'Is property habitable', header_style)
    ws0.write(0, 18, 'Loss Created Date', header_style)
    ws0.write(0, 19, 'Status', header_style)
    ws0.write(0, 20, 'Last Modified By', header_style)
    ws0.write(0, 21, 'Last Modified Date', header_style)

    #valueSize = {}
    #valueSize[row_number, column_num] = len(item)

    row_number = 1
    # Paste the results in the work book
    for row in result_set:
        column_num = 0
        for item in row:  # i.e. for each field in that row
            ws0.write(row_number, column_num, str(item), data_style)  # ,wb.get_sheet(0).cell(0,0).style)  #write excel cell from the cursor at row 1
            column_num = column_num + 1  # increment the column to get the next field

        row_number = row_number + 1

    # for row_index in range(0, 100):
    #     for column_index in range(0, 21):
    #         cwidth = ws0.col(column_index).width
    #         if (len(ws0) * 367) > cwidth:
    #             sheet.col(column_index).width = (
    #                         len(column_data) * 367)  # (Modify column width to match biggest data in that column)

    # Set column width for the sheet
    for i in range(22):
        if i == 16:
           ws0.col(i).width = 256 * 50
        else:
           ws0.col(i).width = 256 * 20
           i = i + 1

    # Set header row height for the sheet
    ws0.row(0).height_mismatch = True
    ws0.row(0).height = 256 * 2

    # Path where the reports are stored
    fileLocation = "C:/Reports/Daily/Stingray/Daily FNOL Report/"
    # Set filenames and subject name for email
    yesterdayDate = inputDate - timedelta(1)
    fileNameDate = yesterdayDate.strftime("%m-%d-%Y")
    subjectDate = yesterdayDate.strftime("%m/%d/%Y")

    # Save the excel file
    fileName = 'FNOL Report ' + fileNameDate + '.xls'
    wb.save(join(fileLocation, fileName))

    # SMTPLIB Code to send email
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email import encoders
    # Define from and to address
    msg = MIMEMultipart()
    recipients = ['csmgmt@relyonanchor.com', 'StingraySupport@relyonanchor.com']
    #recipients = ['StingraySupport@relyonanchor.com']
    #recipients = ['imandapati@relyonanchor.com', 'tbobbili@Relyonanchor.com']
    msg["Subject"] = "FNOL Report - " + subjectDate
    msg["From"] = "stingraydev@relyonanchor.com"
    msg["To"] = "csmgmt@relyonanchor.com; StingraySupport@relyonanchor.com"
    # Add Body
    msg.attach(MIMEText(
        "Good morning,\n\nPlease find attached report for FNOLs created on " + subjectDate + ".\n\nThank you,\nStingray IT", 'plain', 'utf-8'))

    # Add attachment
    attachFile = MIMEBase('application', "octet-stream")
    attachFile.set_payload(open(fileLocation + fileName, 'rb').read())
    encoders.encode_base64(attachFile)
    attachFile.add_header('Content-Disposition', 'attachment; fileName= ' + fileName + '')
    msg.attach(attachFile)
    # Send Email
    s = smtplib.SMTP("relyonanchor-com.mail.protection.outlook.com")
    s.sendmail("stingraydev@relyonanchor.com", recipients, msg.as_string())

    s.quit()

    """
    # Code to Send email from Outlook with File Attachments
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = "Stingray Daily Claims report - " + subjectDate
    newMail.Body = "Good morning,\n\nPlease find attached daily open/closed/reopen claims report from Stingray for " + subjectDate + ".\n\nThanks,\nTeja"
    # newMail.To = "tbobbili@relyonanchor.com"
    newMail.To = "weeklystingrayclaimsreport@relyonanchor.com"
    newMail.CC = "Caity Carpenter <ccarpenter@Relyonanchor.com>"
    # newMail.BCC = "tbobbili@relyonanchor.com"
    attachment1 = ("C:/Users/tbobbili/Desktop/Task/Stingray Daily Claims Report SRD-876/Stingray Daily Claims report - " + fileNameDate + ".xls")
    # attachment2 = ("C:/Users/tbobbili/Desktop/DailyActivity - 12042017.xls")
    # attachment2 = "Path to attachment no. 2"
    newMail.Attachments.Add(attachment1)
    # newMail.Attachments.Add(attachment2)
    newMail.display()
    # newMail.Send()
    print("Email generated successfully!!")
    """


todayDate = date.today()

# weekNo = todayDate.weekday()
#
# holiday_list = ['2018-08-01', '2018-08-11', '2018-09-03', '2018-11-20']
#
# if todayDate.strftime('%Y-%m-%d') in holiday_list or weekNo >= 5:
#     print('Report not run due to company holiday')
# else:
#     print('Running script')

run_report(todayDate)
