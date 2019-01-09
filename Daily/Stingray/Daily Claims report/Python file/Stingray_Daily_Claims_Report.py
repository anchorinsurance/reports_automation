import pyodbc
import xlwt
import win32com.client
from os.path import join
from xlwt import *
from datetime import date, timedelta
import pyodbc
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

def run_report(inputdate):


    # Connection to Dataware House
    cnxn = pyodbc.connect('Driver={SQL Server Native Client 11.0};Server=dwh.relyonanchor.com;Database=AIH_Insurance;Trusted_Connection=yes;')

    #Initiare a Cursor
    cursor = cnxn.cursor()
    # Query for daily balance forward/ renewal notification report for Underwriting taken from SRD-861
    # Begin Query and execution
    q1 = """select distinct CD.CD_ClaimCode as 'Claim Number'\
            , cast(CD.CD_LossReportDate as date) as 'Date Reported'\
            , cast(CD.CD_DOL as date) as 'Date of Loss'\
            , CS.CS_Status as 'Claims Status'\
            , REPLACE(ISNULL(cast(CD.CD_ReOpenDate as date),''),'1900-01-01', '') as 'Date Reopened'\
            , ISNULL(CV1.CV_Name,'') as 'PA / Attorney assigned' \
            , ISNULL(CV2.CV_Name,'') as 'AOB assigned' \
            , SU.sUSR_FirstName + ' ' + SU.sUSR_LastName as 'Adjuster assigned'\
            , case when CC1.CA_ID IS NULL then '' else CC1.CA_Catastrophe end as 'Cat Loss 1'\
            , case when CC2.CA_ID IS NULL then '' else CC2.CA_Catastrophe end as 'Cat Loss 2'\
            , REPLACE(ISNULL(cast(CD_CloseDate as date),''),'1900-01-01', '') as 'Last Close Date'\
            , CCL.CCL_Loss as 'Cause of Loss'\
            , SPR.sPER_Peril as 'Peril'\
            , SF.sFRM_FormType as 'Form Type'\
            , ROUND(SUM(case when CTT_ResTrans = 1 and CTT_IndTrans = 1 then isnull((isnull(CT.CTR_DebitAmt, 0)) - (isnull(CT.CTR_CreditAmt,0)), 0) else 0 end),2) as 'Total Indemnity Reserve Amount'  \
            , ROUND(SUM(case when CTT_ResTrans = 0 and CTT_IndTrans = 1 then isnull((isnull(CT.CTR_DebitAmt, 0)) - (isnull(CT.CTR_CreditAmt,0)), 0) else 0 end),2) as 'Total Indemnity Paid Amount'\
            , ROUND(SUM(case when CTT_ResTrans = 1 and CTT_ExpTrans = 1 then isnull((isnull(CT.CTR_DebitAmt, 0)) - (isnull(CT.CTR_CreditAmt,0)), 0) else 0 end),2) as 'Total Expense Reserve Amount' \
            , ROUND(SUM(case when CTT_ResTrans = 0 and CTT_ExpTrans = 1 then isnull((isnull(CT.CTR_DebitAmt, 0)) - (isnull(CT.CTR_CreditAmt,0)), 0) else 0 end),2) as 'Total Expense Paid Amount'\
            , case when CD_IsClaimOnly = 1 then CO_FirstName + ' ' + CO_LastName else PC_FirstName + ' ' + PC_LastName end as InsuredName\
            , case when CD_IsClaimOnly = 1 then CO_Address1 + ' ' + ISNULL(CO_Address2,'') else PDLI_Address1 + ' ' + ISNULL(PDLI_Address2,'') end as StreetAddress\
            , case when CD_IsClaimOnly = 1 then CO_City else PDLI_City end as City\
            , case when CD_IsClaimOnly = 1 then CO_State else PDLI_State end as State\
            , case when CD_IsClaimOnly = 1 then CO_Zip else PDLI_Zip end as Zip	\
        from Claims_ClaimData CD\
        left join Claims_Vendors CV1 ON CD.CV_ID_PAA = CV1.CV_ID\
        left join Claims_Vendors CV2 ON CD.CV_ID_AOB = CV2.CV_ID\
        left join System_Users SU ON CD.CD_ClaimAdjuster = SU.sUSR_ID\
        left join Claims_Status CS ON CD.CD_StatusID = CS.CS_ID\
        left join Claims_Catastrophes CC1 ON CD.CD_LossCat1 = CC1.CA_ID\
        left join Claims_Catastrophes CC2 ON CD.CD_LossCat2 = CC2.CA_ID\
        left join System_Perils SPR ON CD.CD_PerilID = SPR.sPER_ID\
        left join System_Forms SF ON CD.sFRM_ID = SF.sFRM_ID\
        left join Claims_CauseOfLoss CCL ON CD.CCL_ID = CCL.CCL_ID\
        left JOIN Claims_ClaimsOnly AS CCO ON CCO.CO_ID = CD.PD_ID\
        left JOIN PolicyData AS PD ON PD.PD_PolicyCode = CD.PD_PolicyCode and PD.PD_CurrentRecord=1\
        left JOIN Policy_Contacts AS PC ON PD.PC_ID = PC.PC_ID\
        left JOIN PolicyData_LocationInfo AS PDLI ON PDLI.PD_ID = PD.PD_ID\
        left join Claims_Transactions CT ON CD.CD_ClaimCode = CT.CD_ClaimCode\
        left join Claims_TransactionList as CTL on CTL.CTT_ID = CT.CTT_ID \
        where CD_CurrentRecord = 1 and CD.CD_ClaimCode is not null\
        group by CD.CD_ClaimCode	, cast(CD.CD_LossReportDate as date)	, cast(CD.CD_DOL as date)	, CS.CS_Status	, cast(CD.CD_ReOpenDate as date)\
            , ISNULL(CV1.CV_Name,'')	, ISNULL(CV2.CV_Name,'')	, SU.sUSR_FirstName + ' ' + SU.sUSR_LastName	, CC1.CA_ID	, CC2.CA_ID\
            , CC1.CA_Catastrophe	, CC2.CA_Catastrophe	, cast(CD_CloseDate as date)	, CCL.CCL_Loss	, SPR.sPER_Peril	, SF.sFRM_FormType\
            , CD_IsClaimOnly	, CO_FirstName + ' ' + CO_LastName	, CO_Address1 + ' ' + ISNULL(CO_Address2,'')	, CO_City	, CO_State	, CO_Zip\
            , PC_FirstName + ' ' + PC_LastName	, PDLI_Address1 + ' ' + ISNULL(PDLI_Address2,'')	, PDLI_City	, PDLI_State	, PDLI_Zip\
        order by CD.CD_ClaimCode; """

    # Execute the query
    cursor.execute(q1)

    # Save the results in result_set variable
    result_set = cursor.fetchall()

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
    ws0.write(0, 0, 'Claim Number', style1)
    ws0.write(0, 1, 'Date Reported', style1)
    ws0.write(0, 2, 'Date of Loss', style1)
    ws0.write(0, 3, 'Claim Status', style1)
    ws0.write(0, 4, 'Date Reopened', style1)
    ws0.write(0, 5, 'PA/Attorney assigned', style1)
    ws0.write(0, 6, 'AOB assigned', style1)
    ws0.write(0, 7, 'Adjuster assigned', style1)
    ws0.write(0, 8, 'Cat Loss 1', style1)
    ws0.write(0, 9, 'Cat Loss 2', style1)
    ws0.write(0, 10, 'Last Close Date', style1)
    ws0.write(0, 11, 'Cause of Loss', style1)
    ws0.write(0, 12, 'Peril', style1)
    ws0.write(0, 13, 'Form Type', style1)
    ws0.write(0, 14, 'Total Indemnity Reserve Amount', style1)
    ws0.write(0, 15, 'Total Indemnity Paid Amount', style1)
    ws0.write(0, 16, 'Total Expense Reserve Amount', style1)
    ws0.write(0, 17, 'Total Expense Paid Amount', style1)
    ws0.write(0, 18, 'Insured Name', style1)
    ws0.write(0, 19, 'Street Address', style1)
    ws0.write(0, 20, 'City', style1)
    ws0.write(0, 21, 'State', style1)
    ws0.write(0, 22, 'ZIP', style1)

    row_number = 1
    # Paste the results in the work book
    for row in result_set:
        column_num = 0
        for item in row:                                # i.e. for each field in that row
            ws0.write(row_number, column_num, str(item), style)  # ,wb.get_sheet(0).cell(0,0).style)  #write excel cell from the cursor at row 1
            column_num = column_num + 1  # increment the column to get the next field

        row_number = row_number + 1

    # Set column width for the sheet
    for i in range(21):
        if i == 1:
            ws0.col(i).width = 256 * 20
        else:
            ws0.col(i).width = 256 * 15
        i = i + 1

    # Path where the reports are stored
    test_files = "C:/Reports/Daily/Stingray/Daily Claims report/"

    # Set filenames and subject name for email
    a = inputdate - timedelta(1)
    filename = a.strftime("%m%d%Y")
    subjectname = a.strftime("%m/%d/%Y")

    # Save the excel file
    fname = 'Stingray Daily Claims report - ' + filename + '.xls'
    wb.save(join(test_files, fname))

    # SMTPLIB Code to send email
    # Define from and to address
    msg = MIMEMultipart()
    msg["Subject"] = "Stingray Daily Claims report - " + subjectname
    msg["From"] = "StingrayDev@relyonanchor.com"
    to_address = ['stgdlyclms@relyonanchor.com','stingraysupport@relyonanchor.com','tbobbili@relyonanchor.com','shimpy@relyonanchor.com']
    msg["To"] = "stgdlyclms@relyonanchor.com;"
    msg["CC"] = "stingraysupport@relyonanchor.com; tbobbili@relyonanchor.com; shimpy@relyonanchor.com"
    # Add Body
    msg.attach(MIMEText("Good morning,\n\nPlease find attached daily open/closed/reopen claims report from Stingray for " + subjectname + ".\n\nThank you, Stingray IT",'plain', 'utf-8'))
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
    print ('Report email sent')

def send_email_DWH_notCurrent():
    msg = MIMEMultipart()
    msg["Subject"] = "DWH is not current to send Daily Claims Report"
    msg["From"] = "StingrayDev@relyonanchor.com"
    to_address = ['tbobbili@relyonanchor.com','shimpy@relyonanchor.com', 'ccarpenter@Relyonanchor.com']
    msg["To"] = "tbobbili@relyonanchor.com; Shimpy <shimpy@relyonanchor.com>; Caity Carpenter <ccarpenter@Relyonanchor.com>"
    # Add Body
    msg.attach(MIMEText("Good morning,\n\nNote: DWH is not current with latest data to send daily claims report. Please send the report manually once refresh is complete.\n\nThank you, Stingray IT",'plain', 'utf-8'))
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
    holiday_list = ['2018-11-22', '2018-11-23']

    if todaydate.strftime('%Y-%m-%d') in holiday_list:  # or weekno >= 5:
        print('Report not run due to company holiday')
    else:
        print('Running script')
        DWH_check = check_DWH_isCurrent(todaydate)
        if DWH_check == 1:
            run_report(todaydate)


