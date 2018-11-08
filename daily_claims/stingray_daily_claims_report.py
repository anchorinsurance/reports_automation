from os.path import join
from xlwt import *
from datetime import date, timedelta
import time
import pyodbc
import util

from configparser import ConfigParser

config = ConfigParser()
config.read('./config.ini')

connect_string = "Driver={SQL Server Native Client 11.0};"
connect_string = connect_string + "Server=" + config.get('database', 'db_host') + ";"
connect_string = connect_string + "Database=" + config.get('database', 'db_name') + ";"
if (config.get('database', 'trusted') == 'true'):
    connect_string = connect_string + "Trusted_Connection=yes;"
else:
    connect_string = connect_string + "Uid=" + config.get('database', 'db_user') + ";"
    connect_string = connect_string + "Pwd=" + config.get('database', 'db_password') + ";"

from_address = config.get('email', 'from_address')
file_suffix = (date.today() - timedelta(1)).strftime("%m%d%Y")
subject_suffix = (date.today() - timedelta(1)).strftime("%m/%d/%Y")

output_folder = config.get('folders', 'output_folder')
log_folder = config.get('folders', 'log_folder')
log_file = join("./log", "daily_claims_" + date.today().strftime("%m%d%Y") + ".log")

def run_report(conn, input_date):

    cursor = conn.cursor()
    query = """select distinct CD.CD_ClaimCode as 'Claim Number'\
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

    cursor.execute(query)
    result_set = cursor.fetchall()
    cursor.close()

    # Create an Excel Sheet
    wb = Workbook()
    ws0 = wb.add_sheet('Sheet 1')

    # Set Style_string for headers & rows
    style_row = easyxf("font: bold off, height 220, name Calibri")
    style_header = easyxf("font: bold on, height 220, name Calibri; pattern: pattern solid, fore_colour gray25; " \
                    "borders: top_color black, bottom_color black, right_color black, left_color black,\
                                  left thin, right thin, top thin, bottom thin;")

    # Add headers to the excel workbook
    ws0.write(0, 0, 'Claim Number', style_header)
    ws0.write(0, 1, 'Date Reported', style_header)
    ws0.write(0, 2, 'Date of Loss', style_header)
    ws0.write(0, 3, 'Claim Status', style_header)
    ws0.write(0, 4, 'Date Reopened', style_header)
    ws0.write(0, 5, 'PA/Attorney assigned', style_header)
    ws0.write(0, 6, 'AOB assigned', style_header)
    ws0.write(0, 7, 'Adjuster assigned', style_header)
    ws0.write(0, 8, 'Cat Loss 1', style_header)
    ws0.write(0, 9, 'Cat Loss 2', style_header)
    ws0.write(0, 10, 'Last Close Date', style_header)
    ws0.write(0, 11, 'Cause of Loss', style_header)
    ws0.write(0, 12, 'Peril', style_header)
    ws0.write(0, 13, 'Form Type', style_header)
    ws0.write(0, 14, 'Total Indemnity Reserve Amount', style_header)
    ws0.write(0, 15, 'Total Indemnity Paid Amount', style_header)
    ws0.write(0, 16, 'Total Expense Reserve Amount', style_header)
    ws0.write(0, 17, 'Total Expense Paid Amount', style_header)
    ws0.write(0, 18, 'Insured Name', style_header)
    ws0.write(0, 19, 'Street Address', style_header)
    ws0.write(0, 20, 'City', style_header)
    ws0.write(0, 21, 'State', style_header)
    ws0.write(0, 22, 'ZIP', style_header)

    row_number = 1
    # Paste the results in the work book
    for row in result_set:
        column_num = 0
        for item in row:                                # i.e. for each field in that row
            ws0.write(row_number, column_num, str(item), style_row)  #write excel cell from the cursor at row 1
            column_num = column_num + 1  # increment the column to get the next field

        row_number = row_number + 1

    # Set column width for the sheet
    for i in range(21):
        if i == 1:
            ws0.col(i).width = 256 * 20
        else:
            ws0.col(i).width = 256 * 15
        i = i + 1

    # Save the excel file
    file_name = 'Stingray Daily Claims report - ' + file_suffix + '.xls'
    wb.save(join(output_folder, file_name))

    util.send_mail(from_address,
       ['spepakayala@relyonanchor.com'],
       [],
       "Stingray Daily Claims report - " + subject_suffix,
       "Good morning,\n\n" \
            + "Please find attached daily open/closed/reopen claims report from Stingray for " \
                + subject_suffix + ".\n\nThank you, Stingray IT",
       [join(output_folder, file_name)]
    )
    print ('Report email sent')

def send_email_dwh_not_current(file=None):
    util.send_mail(from_address,
        ['spepakayala@relyonanchor.com'],
        [],
        "DWH is not current to send Daily Claims Report",
        "Good morning,\n\n" \
            + "Note: DWH is not current with latest data to send daily claims report. " \
            + "The process will check in 30 mins and send the report if the DB is refreshed\n\n" \
            + "Thank you, Stingray IT",
        [file]
    )

def is_dwh_current(conn, input_date):

    cursor = conn.cursor()
    query = "select top 1 cast(PD_EntryDate as date) from PolicyData order by 1 desc"
    cursor.execute(query)
    result = cursor.fetchall()
    cursor.close()
    dwh_last_entry_date = result[0][0]

    if dwh_last_entry_date == input_date + timedelta(days=-1):
        return True
    else:
        return False

if __name__ == "__main__":

    today_date = date.today()
    holiday_list = ['2018-11-22', '2018-11-23']

    util.log(log_file, "Process started", False)
    success = False
    while not success:
        try:
            conn = pyodbc.connect(connect_string)
            if not is_dwh_current(conn, today_date):
                util.log(log_file, "DWH is not current", False)
                send_email_dwh_not_current(log_file)
                continue
            run_report(conn, today_date)
            util.log(log_file, "Report sent successfully", False)
            success = True
        except Exception as e:
            util.log(log_file, repr(e), False)
            send_email_dwh_not_current(log_file)
            time.sleep(180)
            success = False
    util.log(log_file, "Process ended", False)